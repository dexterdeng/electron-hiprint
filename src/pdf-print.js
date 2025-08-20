/*
 * @Description: pdf打印
 * @Author: CcSimple
 * @Github: https://github.com/CcSimple
 * @Date: 2023-04-21 16:35:07
 * @LastEditors: CcSimple
 * @LastEditTime: 2025-08-20 12:00:00
 */
const pdfPrint1 = require("pdf-to-printer");
const pdfPrint2 = require("unix-print");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawn } = require("child_process"); // ★ 新增：兜底调用 Adobe
const log = require("../tools/log");
const { store } = require("../tools/utils");
const dayjs = require("dayjs");
const { v7: uuidv7 } = require("uuid");

const printPdfFunction =
  process.platform === "win32" ? pdfPrint1.print : pdfPrint2.print;

const randomStr = () => {
  return Math.random().toString(36).substring(2);
};

// ★ 新增：Windows 下兜底用 Adobe Reader 静默打印
function tryAdobeFallback(pdfPath, printer) {
  return new Promise((resolve, reject) => {
    // 允许从设置里覆盖路径
    const configured = store.get("adobePath");
    // 常见安装路径（优先 64-bit，再 32-bit）
    const candidates = configured
      ? [configured]
      : [
          "C:\\Program Files\\Adobe\\Acrobat\\Acrobat\\Acrobat.exe",
          "C:\\Program Files\\Adobe\\Acrobat Reader\\Reader\\AcroRd32.exe",
          "C:\\Program Files (x86)\\Adobe\\Acrobat Reader\\Reader\\AcroRd32.exe",
        ];
    const exe = candidates.find((p) => {
      try {
        fs.accessSync(p, fs.constants.X_OK);
        return true;
      } catch {
        return false;
      }
    });
    if (!exe) return reject(new Error("未找到 Adobe Reader/Acrobat 可执行文件"));

    log(`fallback: use Adobe to print: ${exe} -> ${printer}`);
    const args = ["/n", "/t", pdfPath, printer];
    const proc = spawn(exe, args, { windowsHide: true });
    let stderr = "";
    proc.stderr.on("data", (d) => (stderr += d.toString()));
    proc.on("error", (e) => reject(e));
    proc.on("close", (code) => {
      if (code === 0) resolve(true);
      else reject(new Error(stderr || `Adobe exit ${code}`));
    });
  });
}

const realPrint = (pdfPath, printer, data, resolve, reject) => {
  if (!fs.existsSync(pdfPath)) {
    reject({ path: pdfPath, msg: "file not found" });
    return;
  }

  if (process.platform === "win32") {
    data = Object.assign({}, data);
    data.printer = printer;
    log("print pdf:" + pdfPath + JSON.stringify(data));
    // 参数见 node_modules/pdf-to-printer/dist/print/print.d.ts
    // pdf打印文档：https://www.sumatrapdfreader.org/docs/Command-line-arguments
    // pdf-to-printer 源码: https://github.com/artiebits/pdf-to-printer
    let pdfOptions = Object.assign(data, { paperSize: data.paperName });
    printPdfFunction(pdfPath, pdfOptions)
      .then(() => {
        resolve();
      })
      .catch(async (e) => {
        log("pdf-to-printer failed, try Adobe fallback: " + (e?.message || e));
        try {
          await tryAdobeFallback(pdfPath, printer); // ★ 新增兜底
          resolve();
        } catch (err) {
          reject(err);
        }
      });
  } else {
    // 参数见 lp 命令 使用方法
    let options = [];
    printPdfFunction(pdfPath, printer, options)
      .then(() => {
        resolve();
      })
      .catch((e) => {
        reject(e);
      });
  }
};

const printPdf = (pdfPath, printer, data) => {
  return new Promise((resolve, reject) => {
    try {
      if (typeof pdfPath !== "string") {
        reject("pdfPath must be a string");
      }
      if (/^https?:\/\/.+/.test(pdfPath)) {
        const client = pdfPath.startsWith("https") ? require("https") : require("http");
        client
          .get(pdfPath, (res) => {
            const toSavePath = path.join(
              store.get("pdfPath") || os.tmpdir(),
              "url_pdf",
              dayjs().format(`YYYY_MM_DD HH_mm_ss_`) + `${uuidv7()}.pdf`,
            );
            // 确保目录存在
            fs.mkdirSync(path.dirname(toSavePath), { recursive: true });
            const file = fs.createWriteStream(toSavePath);
            res.pipe(file);
            file.on("finish", () => {
              file.close();
              log("file downloaded:" + toSavePath);
              realPrint(toSavePath, printer, data, resolve, reject);
            });
          })
          .on("error", (err) => {
            log("download pdf error:" + err?.message);
            reject(err);
          });
        return;
      }
      realPrint(pdfPath, printer, data, resolve, reject);
    } catch (error) {
      log("print error:" + error?.message);
      reject(error);
    }
  });
};

/**
 * @description: 打印Blob类型的PDF数据
 * @param {Blob|Uint8Array|Buffer} pdfBlob PDF的二进制数据
 * @param {string} printer 打印机名称
 * @param {object} data 打印参数
 * @return {Promise}
 */
const printPdfBlob = (pdfBlob, printer, data) => {
  return new Promise((resolve, reject) => {
    try {
      // 验证blob数据 实际是 Uint8Array（Node18+ 也有 Blob）
      if (
        !pdfBlob ||
        !(
          (typeof Blob !== "undefined" && pdfBlob instanceof Blob) ||
          pdfBlob instanceof Uint8Array ||
          Buffer.isBuffer(pdfBlob)
        )
      ) {
        reject(new Error("pdfBlob must be a Blob, Uint8Array, or Buffer"));
        return;
      }

      // 生成临时文件路径
      const toSavePath = path.join(
        store.get("pdfPath") || os.tmpdir(),
        "blob_pdf",
        dayjs().format(`YYYY_MM_DD HH_mm_ss_`) + `${uuidv7()}.pdf`,
      );

      // 确保目录存在
      fs.mkdirSync(path.dirname(toSavePath), { recursive: true });

      // Uint8Array / Blob → Buffer
      const toBuffer = (blobOrU8) =>
        Buffer.isBuffer(blobOrU8)
          ? blobOrU8
          : typeof Blob !== "undefined" && blobOrU8 instanceof Blob
          ? Buffer.from(new Uint8Array(blobOrU8.buffer || await blobOrU8.arrayBuffer()))
          : Buffer.from(blobOrU8);

      Promise.resolve(toBuffer(pdfBlob))
        .then((buffer) => {
          fs.writeFile(toSavePath, buffer, (err) => {
            if (err) {
              log("save blob pdf error:" + err?.message);
              reject(err);
              return;
            }
            log("blob pdf saved:" + toSavePath);
            realPrint(toSavePath, printer, data, resolve, reject);
          });
        })
        .catch((err) => {
          reject(err);
        });
    } catch (error) {
      log("print blob error:" + error?.message);
      reject(error);
    }
  });
};

module.exports = {
  printPdf,
  printPdfBlob,
};

