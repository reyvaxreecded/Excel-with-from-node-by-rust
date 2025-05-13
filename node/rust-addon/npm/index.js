const { existsSync, readFileSync } = require("fs");
const { join } = require("path");

const { platform, arch } = process;

let nativeBinding = null;
let localFileExisted = false;
let loadError = null;

function isMusl() {
  // Pour Node 10
  if (!process.report || typeof process.report.getReport !== "function") {
    try {
      return readFileSync("/proc/self/maps", "utf8").includes("ld-musl");
    } catch (e) {
      return false;
    }
  } else {
    const { glibcVersionRuntime } = process.report.getReport().header;
    return !glibcVersionRuntime;
  }
}

switch (platform) {
  case "win32":
    switch (arch) {
      case "x64":
        localFileExisted = existsSync(
          join(__dirname, "excel_manager.win32-x64-msvc.node")
        );
        try {
          if (localFileExisted) {
            nativeBinding = require("./excel_manager.win32-x64-msvc.node");
          } else {
            nativeBinding = require("excel_manager-win32-x64-msvc");
          }
        } catch (e) {
          loadError = e;
        }
        break;
      case "ia32":
        localFileExisted = existsSync(
          join(__dirname, "excel_manager.win32-ia32-msvc.node")
        );
        try {
          if (localFileExisted) {
            nativeBinding = require("./excel_manager.win32-ia32-msvc.node");
          } else {
            nativeBinding = require("excel_manager-win32-ia32-msvc");
          }
        } catch (e) {
          loadError = e;
        }
        break;
      case "arm64":
        localFileExisted = existsSync(
          join(__dirname, "excel_manager.win32-arm64-msvc.node")
        );
        try {
          if (localFileExisted) {
            nativeBinding = require("./excel_manager.win32-arm64-msvc.node");
          } else {
            nativeBinding = require("excel_manager-win32-arm64-msvc");
          }
        } catch (e) {
          loadError = e;
        }
        break;
      default:
        throw new Error(`Unsupported architecture on Windows: ${arch}`);
    }
    break;
  case "darwin":
    switch (arch) {
      case "x64":
        localFileExisted = existsSync(
          join(__dirname, "excel_manager.darwin-x64.node")
        );
        try {
          if (localFileExisted) {
            nativeBinding = require("./excel_manager.darwin-x64.node");
          } else {
            nativeBinding = require("excel_manager-darwin-x64");
          }
        } catch (e) {
          loadError = e;
        }
        break;
      case "arm64":
        localFileExisted = existsSync(
          join(__dirname, "excel_manager.darwin-arm64.node")
        );
        try {
          if (localFileExisted) {
            nativeBinding = require("./excel_manager.darwin-arm64.node");
          } else {
            nativeBinding = require("excel_manager-darwin-arm64");
          }
        } catch (e) {
          loadError = e;
        }
        break;
      default:
        throw new Error(`Unsupported architecture on macOS: ${arch}`);
    }
    break;
  case "linux":
    switch (arch) {
      case "x64":
        if (isMusl()) {
          localFileExisted = existsSync(
            join(__dirname, "excel_manager.linux-x64-musl.node")
          );
          try {
            if (localFileExisted) {
              nativeBinding = require("./excel_manager.linux-x64-musl.node");
            } else {
              nativeBinding = require("excel_manager-linux-x64-musl");
            }
          } catch (e) {
            loadError = e;
          }
        } else {
          localFileExisted = existsSync(
            join(__dirname, "excel_manager.linux-x64-gnu.node")
          );
          try {
            if (localFileExisted) {
              nativeBinding = require("./excel_manager.linux-x64-gnu.node");
            } else {
              nativeBinding = require("excel_manager-linux-x64-gnu");
            }
          } catch (e) {
            loadError = e;
          }
        }
        break;
      case "arm64":
        if (isMusl()) {
          localFileExisted = existsSync(
            join(__dirname, "excel_manager.linux-arm64-musl.node")
          );
          try {
            if (localFileExisted) {
              nativeBinding = require("./excel_manager.linux-arm64-musl.node");
            } else {
              nativeBinding = require("excel_manager-linux-arm64-musl");
            }
          } catch (e) {
            loadError = e;
          }
        } else {
          localFileExisted = existsSync(
            join(__dirname, "excel_manager.linux-arm64-gnu.node")
          );
          try {
            if (localFileExisted) {
              nativeBinding = require("./excel_manager.linux-arm64-gnu.node");
            } else {
              nativeBinding = require("excel_manager-linux-arm64-gnu");
            }
          } catch (e) {
            loadError = e;
          }
        }
        break;
      default:
        throw new Error(`Unsupported architecture on Linux: ${arch}`);
    }
    break;
  default:
    throw new Error(`Unsupported OS: ${platform}, architecture: ${arch}`);
}

if (!nativeBinding) {
  if (loadError) {
    throw loadError;
  }
  throw new Error(`Failed to load native binding`);
}

const { read_excel_file, upsert_row } = nativeBinding;

module.exports.read_excel_file = read_excel_file;
module.exports.upsert_row = upsert_row;
