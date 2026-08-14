const { app, BrowserWindow, dialog } = require("electron");
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");
const net = require("net");

let mainWindow = null;
let backendProcess = null;

function findFreePort(startPort = 8765) {
  return new Promise((resolve, reject) => {
    let port = startPort;

    function tryPort() {
      const server = net.createServer();
      server.once("error", () => {
        port += 1;
        if (port > startPort + 50) {
          reject(new Error("没有找到可用端口。"));
          return;
        }
        tryPort();
      });
      server.once("listening", () => {
        server.close(() => resolve(port));
      });
      server.listen(port, "127.0.0.1");
    }

    tryPort();
  });
}

function backendPath() {
  if (app.isPackaged) {
    const backendName = process.platform === "win32" ? "直播投放处理器后端.exe" : "直播投放处理器后端";
    return path.join(process.resourcesPath, "backend", "直播投放处理器后端", backendName);
  }
  return path.join(__dirname, "..", "web_app.py");
}

async function startBackend() {
  const port = await findFreePort();
  const env = {
    ...process.env,
    LIVE_ADS_OCR_DIR: path.join(app.getPath("userData"), "ocr"),
    NO_BROWSER: "1",
    PORT: String(port),
  };

  if (app.isPackaged) {
    const exePath = backendPath();
    if (!fs.existsSync(exePath)) {
      throw new Error(`找不到后端程序：${exePath}`);
    }
    backendProcess = spawn(exePath, [], {
      env,
      windowsHide: true,
      stdio: "ignore",
    });
  } else {
    backendProcess = spawn("python", [backendPath()], {
      env,
      windowsHide: true,
      stdio: "ignore",
    });
  }

  backendProcess.on("exit", () => {
    backendProcess = null;
  });

  return `http://127.0.0.1:${port}`;
}

function waitForServer(url, timeoutMs = 15000) {
  const startedAt = Date.now();

  return new Promise((resolve, reject) => {
    function check() {
      fetch(url)
        .then((response) => {
          if (response.ok) {
            resolve();
            return;
          }
          retry();
        })
        .catch(retry);
    }

    function retry() {
      if (Date.now() - startedAt > timeoutMs) {
        reject(new Error("后端服务启动超时。"));
        return;
      }
      setTimeout(check, 250);
    }

    check();
  });
}

async function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 1080,
    minHeight: 720,
    title: "",
    titleBarStyle: process.platform === "darwin" ? "hiddenInset" : "default",
    trafficLightPosition: { x: 16, y: 16 },
    backgroundColor: "#faf9f5",
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  try {
    const url = await startBackend();
    await waitForServer(url);
    await mainWindow.loadURL(url);
  } catch (error) {
    dialog.showErrorBox("直播投放处理器启动失败", error.message);
    app.quit();
  }
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  app.quit();
});

app.on("before-quit", () => {
  if (backendProcess) {
    backendProcess.kill();
    backendProcess = null;
  }
});
