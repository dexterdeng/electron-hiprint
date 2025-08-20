/*
 * @Description: pdf打印
 * @Author: CcSimple
 * @Github: https://github.com/CcSimple
 * @Date: 2023-04-21 16:35:07
 * @LastEditors: CcSimple
 * @LastEditTime: 2025-08-20 18:30:00
 */
const pdfPrint1 = require("pdf-to-printer");
const pdfPrint2 = require("unix-print");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawn } = require("child_process");
const log = require("../tools/log");
const { store } = require("../tools/utils");
const dayjs = require("dayjs");
const { v7: uuidv7 } = require("uuid");

const printPdfFunction =
  process.platform === "win32" ? pdfPrint1.print : pdfPrint2.print;

const randomStr = () => Math.random().toString(36).substring(2);

/** -----------------------
 * Windows 外部引擎封装
 * ----------------------*/

/** Adobe Reader / Acrobat 静默打印 */
function tryAdobe(pdfPath, printer) {
  return new Promise((resolve, reject) => {
    // 允许从设置覆盖路径
    const configured = store.get("adobePath");
    const candidates = configured
      ? [configured]
      : [
          // Acrobat（完整版）
          "C:\\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
          // Acrobat（完整版）
          "C:\\Program Files\\Adobe\\Acrobat\\Acrobat\\Acrobat.exe",
          // Reader 64-bit
          "C:\\Program Files\\Adobe\\Acrobat Reader\\Reader\\AcroRd32.exe",
          // Reader 32-bit
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

    log(`win pdf engine: Adobe -> ${exe} @ ${printer}`);
    const args = ["/n", "/t", pdfPath, printer];
    const p = spawn(exe, args, { windowsHide: true });
    let stderr = "";
    p.stderr.on("data", (d) => (stderr += d.toString()));
    p.on("error", reject);
    p.on("close", (code) => {
      if (code === 0) resolve(true);
      else reject(new Error(stderr || `Adobe exit ${code}`));
    });
  });
}

/** SumatraPDF 静默打印（备选） */
function trySumatra(pdfPath, printer) {
  return new Promise((resolve, reject) => {
    const configured = store.get("sumatraPath");
    const candidates = configured
      ? [configured]
      : [
          "C:\\Program Files\\SumatraPDF\\SumatraPDF.exe",
          "C:\\Program Files (x86)\\SumatraPDF\\SumatraPDF.exe",
        ];
    const exe = candidates.find((p) => {
      try {
        fs.accessSync(p, fs.constants.X_OK);
        return true;
      } catch {
        return false;
      }
    });
    if (!exe) return reject(new Error("未找到 SumatraPDF 可执行文件"));

    log(`win pdf engine: Sumatra -> ${exe} @ ${printer}`);
    const args = ["-print-to", printer, pdfPath];
    const p = spawn(exe, args, { windowsHide: true });
    let stderr = "";
    p.stderr.on("data", (d) => (stderr += d.toString()));
    p.on("error", reject);
    p.on("close", (code) => {
      if (code === 0) resolve(true);
      else reject(new Error(stderr || `Sumatra exit ${code}`));
    });
  });
}

/** system: pdf-to-printer（依赖系统默认 PDF 处理器） */
function trySystemPdfToPrinter(pdfPath, printer, data) {
  const pdfOptions = Object.assign({}, data, {
    printer,
    paperSize: data?.paperName,
  });
  log(`win pdf engine: system(pdf-to-printer) @ ${printer}`);
  return printPdfFunction(pdfPath, pdfOptions);
}

/** -----------------------
 * 实际打印流程
 * ----------------------*/
const realPrint = (pdfPath, printer, data, resolve, reject) => {
  if (!fs.existsSync(pdfPath)) {
    reject({ path: pdfPath, msg: "file not found" });
    return;
  }

  if (process.platform === "win32") {
    // Windows 顺序：优先 Adobe → 再 System → 最后 Sumatra
    // 可通过设置覆盖：winPdfEngine = adobe | sumatra | system
    const engine = (store.get("winPdfEngine") || "adobe").toLowerCase();

    const chainByEngine = {
      adobe: [() => tryAdobe(pdfPath, printer), () => trySystemPdfToPrinter(pdfPath, printer, data), () => trySumatra(pdfPath, printer)],
      system: [() => trySystemPdfToPrinter(pdfPath, printer, data), () => tryAdobe(pdfPath, printer), () => trySumatra(pdfPath, printer)],
      sumatra: [() => trySumatra(pdfPath, printer), () => trySystemPdfToPrinter(pdfPath, printer, data), () => tryAdobe(pdfPath, printer)],
    };
    const chain = chainByEngine[engine] || chainByEngine.adobe;

    (async () => {
      for (const step of chain) {
        try {
          await step();
          resolve();
          return;
        } catch (e) {
          log("win pdf engine step failed: " + (e?.message || e));
        }
      }
      reject(new Error("所有 Windows PDF 打印引擎均失败"));
    })();
  } else {
    // 参数见 lp 命令 使用方法（macOS / Linux）
    let options = [];
    printPdfFunction(pdfPath, printer, options)
      .then(() => resolve())
      .catch((e) => reject(e));
  }
};

/**
 * @description: 打印PDF（本地路径或 URL）
 * @param {string} pdfPath  本地路径或 http/https URL
 * @param {string} printer  打印机名称
 * @param {object} data     其它打印参数
 * @return {Promise}
 */
const printPdf = (pdfPath, printer, data) => {
  return new Promise((resolve, reject) => {
    try {
      if (typeof pdfPath !== "string") {
        reject("pdfPath must be a string");
        return;
      }
      // URL：先下载到临时目录
      if (/^https?:\/\/.+/.test(pdfPath)) {
        const client =
          pdfPath.startsWith("https") ? require("https") : require("http");
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
      // 本地文件
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
      // 验证 blob 数据（Node18+ 也可能存在 Blob）
      const isValid =
        (typeof Blob !== "undefined" && pdfBlob instanceof Blob) ||
        pdfBlob instanceof Uint8Array ||
        Buffer.isBuffer(pdfBlob);
      if (!isValid) {
        reject(new Error("pdfBlob must be a Blob, Uint8Array, or Buffer"));
        return;
      }

      // 生成临时文件路径
      const toSavePath = path.join(
        store.get("pdfPath") || os.tmpdir(),
        "blob_pdf",
        dayjs().format(`YYYY_MM_DD HH_mm_ss_`) + `${uuidv7()}.pdf`,
      );
      fs.mkdirSync(path.dirname(toSavePath), { recursive: true });

      // 转 Buffer（避免在三元里用 await 的语法错误）
      const toBuffer = async (blobOrU8) => {
        if (Buffer.isBuffer(blobOrU8)) return blobOrU8;
        if (typeof Blob !== "undefined" && blobOrU8 instanceof Blob) {
          const ab = await blobOrU8.arrayBuffer();
          return Buffer.from(new Uint8Array(ab));
        }
        return Buffer.from(blobOrU8);
      };

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
        .catch((err) => reject(err));
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

