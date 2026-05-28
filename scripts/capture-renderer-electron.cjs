const { app, BrowserWindow } = require('electron');
const fs = require('node:fs/promises');
const path = require('node:path');

async function main() {
  await app.whenReady();
  const win = new BrowserWindow({
    show: false,
    width: 1440,
    height: 920,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  await win.loadURL('http://127.0.0.1:4173');
  await new Promise((resolve) => setTimeout(resolve, 800));
  const projectItems = await win.webContents.executeJavaScript(
    'document.querySelectorAll(".project-item").length'
  );
  win.webContents.sendInputEvent({ type: 'mouseDown', x: 390, y: 224, button: 'left', clickCount: 1 });
  win.webContents.sendInputEvent({ type: 'mouseUp', x: 390, y: 224, button: 'left', clickCount: 1 });
  await new Promise((resolve) => setTimeout(resolve, 800));
  const firstHeading = await win.webContents.executeJavaScript(
    'document.querySelector("h1")?.textContent || ""'
  );
  const saveButtons = await win.webContents.executeJavaScript(
    '[...document.querySelectorAll("button")].filter((button) => button.textContent.includes("Salva")).length'
  );
  const image = await win.webContents.capturePage();
  const outputPath = path.resolve('outputs', 'renderer-preview.png');
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, image.toPNG());
  console.log(JSON.stringify({ firstHeading, projectItems, saveButtons, outputPath }, null, 2));
  win.destroy();
  app.quit();
}

main().catch((error) => {
  console.error(error);
  app.quit();
  process.exit(1);
});
