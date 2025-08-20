"use strict";

const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const os = require("os");
const fs = require("fs");
const { pathToFileURL } = require("url"); // 规范 file://
const { printPdf, printPdfBlob } = require("./pdf-print");
const log = require("../tools/log");
const { store, getCurrentPrintStatusByName } = require("../tools/utils");
const db = require("../tools/database");
const dayjs = require("dayjs");
const { v7: uuidv7 } = require("uuid");

// ★ Windows 上很多机器拿不到“状态”，而且会触发外部可执行程序（可能依赖 .NET 3.5）。
//   我们不因“状态未知/异常”阻塞打印，默认继续打印。
const IGNORE_STATUS_ON_WIN32 = true;

function safeGetStatusMsg(printerName) {
  try {
    const info = getCurrentPrintStatusByName(printerName);
    return (info && info.StatusMsg) || "未知状态";
  } catch (e) {
    log(`safeGetStatusMsg error: ${e?.message || e}`);
    return "状态不可用";
  }
}

/**
 * @description: 创建打印窗口
 * @return {BrowserWindow} PRINT_WINDOW 打印窗口
 */
async function createPrintWindow() {
  const windowOptions = {
    width: 100,
    height: 100,
    show: false,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration: true,
    },
    backgroundColor: "#fff",
  };

  PRINT_WINDOW = new BrowserWindow(windowOptions);

  // 规范 file:/// URL
  const printHtml = pathToFileURL(
    path.join(app.getAppPath(), "assets", "print.html"),
  ).toString();
  PRINT_WINDOW.webContents.loadURL(printHtml);

  // init events
  initPrintEvent();

  return PRINT_WINDOW;
}

/**
 * @description: 绑定打印窗口事件
 * @return {Void}
 */
function initPrintEvent() {
  ipcMain.on("do", async (event, data) => {
    let socket = null;
    if (data.clientType === "local") {
      socket = SOCKET_SERVER.sockets.sockets.get(data.socketId);
    } else {
      socket = SOCKET_CLIENT;
    }

    // 取打印机列表
    const printers = await PRINT_WINDOW.webContents.getPrintersAsync();
    let defaultPrinter = data.printer || store.get("defaultPrinter", "");
    let targetPrinter;

    // 选默认机
    printers.forEach((p) => {
      if (p.isDefault && (!defaultPrinter || defaultPrinter === "")) {
        defaultPrinter = p.name;
      }
      if (p.name === defaultPrinter) targetPrinter = p;
    });

    // 如果没选到，允许继续（让系统自己决定），但记录日志
    if (!targetPrinter && defaultPrinter) {
      log(`warn: 指定的打印机未在列表中找到：${defaultPrinter}，将尝试继续打印（可能走默认机）。`);
    }

    // —— 原来这里会强依赖“状态码”并在 Windows 上设为非 0 就判“异常”。
    // 我们调整策略：仅在非 Windows 或明确要求时拦截；默认 Windows 放行。
    let printerError = false;
    if (targetPrinter) {
      if (process.platform === "win32") {
        // Windows：除非用户关闭忽略，否则不拦截
        if (!IGNORE_STATUS_ON_WIN32) {
          printerError = targetPrinter.status != 0;
        }
      } else {
        // mac/linux：沿用原来的 3
        printerError = targetPrinter.status != 3;
      }
    }

    if (printerError) {
      const msg = safeGetStatusMsg(defaultPrinter);
      log(
        `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
          data.templateId
        }】 打印失败，打印机异常：${defaultPrinter}，状态：${msg}`,
      );
      socket &&
        socket.emit("error", {
          msg: defaultPrinter + "打印机异常：" + msg,
          templateId: data.templateId,
          replyId: data.replyId,
        });
      if (data.taskId) {
        PRINT_RUNNER_DONE[data.taskId]();
        delete PRINT_RUNNER_DONE[data.taskId];
      }
      MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
      return;
    }

    const deviceName = defaultPrinter; // 可能为空字符串 → 交给系统默认机
    const logPrintResult = (status, errorMessage = "") => {
      db.run(
        `INSERT INTO print_logs (socketId, clientType, printer, templateId, data, pageNum, status, rePrintAble, errorMessage) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
        [
          socket?.id,
          data.clientType,
          deviceName,
          data.templateId,
          JSON.stringify(data),
          data.pageNum,
          status,
          data.rePrintAble ?? 1,
          errorMessage,
        ],
        (err) => {
          if (err) console.error("Failed to log print result", err);
        },
      );
    };

    // ====== 分支 1：type = "pdf"（把当前页面渲染为 PDF 再打）======
    const isPdf = data.type && `${data.type}`.toLowerCase() === "pdf";
    if (isPdf) {
      const pdfPath = path.join(
        store.get("pdfPath") || os.tmpdir(),
        "hiprint",
        dayjs().format(`YYYY_MM_DD HH_mm_ss_`) + `${uuidv7()}.pdf`,
      );
      fs.mkdirSync(path.dirname(pdfPath), { recursive: true });

      PRINT_WINDOW.webContents
        .printToPDF({
          landscape: data.landscape ?? false,
          displayHeaderFooter: data.displayHeaderFooter ?? false,
          printBackground: data.printBackground ?? true,
          scale: data.scale ?? 1,
          pageSize: data.pageSize,
          margins: data.margins ?? { marginType: "none" },
          pageRanges: data.pageRanges,
          headerTemplate: data.headerTemplate,
          footerTemplate: data.footerTemplate,
          preferCSSPageSize: data.preferCSSPageSize ?? false,
        })
        .then((pdfData) => {
          fs.writeFileSync(pdfPath, pdfData);
          return printPdf(pdfPath, deviceName, data);
        })
        .then(() => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印成功，类型：PDF，打印机：${deviceName}，页数：${
              data.pageNum
            }`,
          );
          if (socket) {
            const result = { msg: "打印成功", templateId: data.templateId, replyId: data.replyId };
            socket.emit("successs", result);
            socket.emit("success", result);
          }
          logPrintResult("success");
        })
        .catch((err) => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印失败，类型：PDF，打印机：${deviceName}，原因：${err?.message || err}`,
          );
          socket &&
            socket.emit("error", {
              msg: "打印失败: " + (err?.message || err),
              templateId: data.templateId,
              replyId: data.replyId,
            });
          logPrintResult("failed", err?.message || String(err));
        })
        .finally(() => {
          if (data.taskId) {
            PRINT_RUNNER_DONE[data.taskId]();
            delete PRINT_RUNNER_DONE[data.taskId];
          }
          MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
        });
      return;
    }

    // ====== 分支 2：type = "url_pdf"（支持 pdf_path 和 pdf_url）======
    const isUrlPdf = data.type && `${data.type}`.toLowerCase() === "url_pdf";
    if (isUrlPdf) {
      const urlOrPath = data.pdf_path || data.pdf_url;
      printPdf(urlOrPath, deviceName, data)
        .then(() => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印成功，类型：URL_PDF，打印机：${deviceName}，页数：${
              data.pageNum
            }`,
          );
          if (socket) {
            const ok = { msg: "打印成功", templateId: data.templateId, replyId: data.replyId };
            socket.emit("successs", ok);
            socket.emit("success", ok);
          }
          logPrintResult("success");
        })
        .catch((err) => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印失败，类型：URL_PDF，打印机：${deviceName}，原因：${err?.message || err}`,
          );
          socket &&
            socket.emit("error", {
              msg: "打印失败: " + (err?.message || err),
              templateId: data.templateId,
              replyId: data.replyId,
            });
          logPrintResult("failed", err?.message || String(err));
        })
        .finally(() => {
          if (data.taskId) {
            PRINT_RUNNER_DONE[data.taskId]();
            delete PRINT_RUNNER_DONE[data.taskId];
          }
          MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
        });
      return;
    }

    // ====== 分支 3：type = "blob_pdf" ======
    const isBlobPdf = data.type && `${data.type}`.toLowerCase() === "blob_pdf";
    if (isBlobPdf) {
      if (!data.pdf_blob) {
        const errorMsg = "blob_pdf 缺少 pdf_blob 参数";
        log(`${socket?.id} 模板【${data.templateId}】 打印失败：${errorMsg}`);
        socket &&
          socket.emit("error", {
            msg: errorMsg,
            templateId: data.templateId,
            replyId: data.replyId,
          });
        logPrintResult("failed", errorMsg);
        if (data.taskId) {
          PRINT_RUNNER_DONE[data.taskId]();
          delete PRINT_RUNNER_DONE[data.taskId];
        }
        MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
        return;
      }
      const pdfBlob = data.pdf_blob;
      delete data.pdf_blob;
      printPdfBlob(pdfBlob, deviceName, data)
        .then(() => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印成功，类型：BLOB_PDF，打印机：${deviceName}，页数：${
              data.pageNum
            }`,
          );
          if (socket) {
            const ok = { msg: "打印成功", templateId: data.templateId, replyId: data.replyId };
            socket.emit("successs", ok);
            socket.emit("success", ok);
          }
          logPrintResult("success");
        })
        .catch((err) => {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印失败，类型：BLOB_PDF，打印机：${deviceName}，原因：${err?.message || err}`,
          );
          socket &&
            socket.emit("error", {
              msg: "打印失败: " + (err?.message || err),
              templateId: data.templateId,
              replyId: data.replyId,
            });
          logPrintResult("failed", err?.message || String(err));
        })
        .finally(() => {
          if (data.taskId) {
            PRINT_RUNNER_DONE[data.taskId]();
            delete PRINT_RUNNER_DONE[data.taskId];
          }
          MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
        });
      return;
    }

    // ====== 分支 4：HTML 直接打印 ======
    PRINT_WINDOW.webContents.print(
      {
        silent: data.silent ?? true,
        printBackground: data.printBackground ?? true,
        deviceName: deviceName, // 允许为空 → 系统默认机
        color: data.color ?? true,
        margins: data.margins ?? { marginType: "none" },
        landscape: data.landscape ?? false,
        scaleFactor: data.scaleFactor ?? 100,
        pagesPerSheet: data.pagesPerSheet ?? 1,
        collate: data.collate ?? true,
        copies: data.copies ?? 1,
        pageRanges: data.pageRanges ?? {},
        duplexMode: data.duplexMode,
        dpi: data.dpi ?? 300,
        header: data.header,
        footer: data.footer,
        pageSize: data.pageSize,
      },
      (success, failureReason) => {
        if (success) {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印成功，类型：HTML，打印机：${deviceName}，页数：${
              data.pageNum
            }`,
          );
          logPrintResult("success");
        } else {
          log(
            `${data.replyId ? "中转服务" : "插件端"} ${socket?.id} 模板【${
              data.templateId
            }】 打印失败，类型：HTML，打印机：${deviceName}，原因：${failureReason}`,
          );
          logPrintResult("failed", failureReason);
        }
        if (socket) {
          if (success) {
            const ok = { msg: "打印成功", templateId: data.templateId, replyId: data.replyId };
            socket.emit("successs", ok);
            socket.emit("success", ok);
          } else {
            socket.emit("error", {
              msg: failureReason,
              templateId: data.templateId,
              replyId: data.replyId,
            });
          }
        }
        if (data.taskId) {
          PRINT_RUNNER_DONE[data.taskId]();
          delete PRINT_RUNNER_DONE[data.taskId];
        }
        MAIN_WINDOW.webContents.send("printTask", PRINT_RUNNER.isBusy());
      },
    );
  });
}

module.exports = async () => {
  await createPrintWindow();
};
