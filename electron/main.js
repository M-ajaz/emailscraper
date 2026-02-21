const { app, BrowserWindow, Tray, Menu, nativeImage, shell } = require('electron')
const { spawn } = require('child_process')
const path = require('path')
const fs = require('fs')
const http = require('http')
const net = require('net')
const isDev = require('electron-is-dev')

// ─── Packaged App Paths ─────────────────────────────────────────────
const RESOURCES_PATH = app.isPackaged
  ? process.resourcesPath
  : path.join(__dirname, '..')

// ─── Global Variables ────────────────────────────────────────────────
let mainWindow = null
let splashWindow = null
let tray = null
let backendProcess = null

const BACKEND_URL = 'http://localhost:8000'
const FRONTEND_URL = app.isPackaged
  ? `file://${path.join(RESOURCES_PATH, 'frontend', 'index.html')}`
  : (isDev ? 'http://localhost:5173' : 'http://localhost:5173')

// ─── Port Check ─────────────────────────────────────────────────────
function isPortInUse(port) {
  return new Promise((resolve) => {
    const tester = net.createServer()
      .once('error', () => resolve(true))
      .once('listening', () => { tester.close(); resolve(false) })
      .listen(port)
  })
}

// ─── Backend Process ─────────────────────────────────────────────────
async function startBackend() {
  const portBusy = await isPortInUse(8000)
  if (portBusy) {
    console.log('[Backend] Port 8000 already in use — assuming backend already running')
    return
  }

  const backendPath = app.isPackaged
    ? path.join(RESOURCES_PATH, 'backend')
    : path.join(__dirname, '..', 'backend')

  const envPath = path.join(backendPath, '.env')
  const templatePath = path.join(backendPath, '.env.template')
  if (!fs.existsSync(envPath) && fs.existsSync(templatePath)) {
    fs.copyFileSync(templatePath, envPath)
    console.log('[Setup] Created .env from template')
  }

  if (app.isPackaged) {
    // Use compiled PyInstaller binary
    const ext = process.platform === 'win32' ? '.exe' : ''
    const binaryPath = path.join(
      RESOURCES_PATH,
      'backend-dist',
      'mailscraper-backend',
      `mailscraper-backend${ext}`
    )
    backendProcess = spawn(binaryPath, [], {
      cwd: backendPath,
      windowsHide: true,
      env: { ...process.env, PORT: '8000' },
    })
  } else {
    // Development: use system Python
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3'
    backendProcess = spawn(
      pythonCmd,
      ['-m', 'uvicorn', 'main:app', '--port', '8000', '--reload'],
      {
        cwd: backendPath,
        windowsHide: true,
      }
    )
  }

  let startTime = Date.now()

  backendProcess.stdout.on('data', (d) => console.log('[Backend]', d.toString()))
  backendProcess.stderr.on('data', (d) => console.error('[Backend]', d.toString()))
  backendProcess.on('exit', (code) => {
    console.log('[Backend] exited with code', code)
    if (Date.now() - startTime < 5000) {
      const { dialog } = require('electron')
      dialog.showErrorBox(
        'Backend Failed to Start',
        'The Mail Scraper backend crashed on startup.\n\nPossible causes:\n• Python not installed\n• Missing packages (run INSTALL_WINDOWS.bat)\n• Port 8000 already in use\n• .env file missing\n\nCheck the logs for details.'
      )
    }
    backendProcess = null
  })
}

// ─── Wait for Backend ────────────────────────────────────────────────
function waitForBackend(retries = 30) {
  return new Promise((resolve, reject) => {
    const check = (n) => {
      http.get(`${BACKEND_URL}/docs`, (res) => {
        if (res.statusCode === 200) resolve()
        else if (n > 0) setTimeout(() => check(n - 1), 1500)
        else reject(new Error('Backend failed to start'))
      }).on('error', () => {
        if (n > 0) setTimeout(() => check(n - 1), 1500)
        else reject(new Error('Backend not responding'))
      })
    }
    // Wait 3 seconds before first check to let Python start up
    setTimeout(() => check(retries), 3000)
  })
}

// ─── Splash Screen ───────────────────────────────────────────────────
function createSplashWindow() {
  splashWindow = new BrowserWindow({
    width: 400,
    height: 300,
    resizable: false,
    frame: false,
    alwaysOnTop: true,
    transparent: false,
    backgroundColor: '#0f1117',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
  })

  splashWindow.loadFile(path.join(__dirname, 'splash.html'))
  splashWindow.center()
}

// ─── Main Window ─────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
    },
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    icon: path.join(__dirname, 'assets', 'icon.png'),
    show: false,
  })

  if (app.isPackaged) {
    mainWindow.loadFile(path.join(RESOURCES_PATH, 'frontend', 'index.html'))
  } else {
    mainWindow.loadURL('http://localhost:5173')
  }

  mainWindow.once('ready-to-show', () => {
    // show is called from app.whenReady after splash closes
  })

  mainWindow.on('close', (e) => {
    if (!app.isQuiting) {
      e.preventDefault()
      mainWindow.hide()
    }
  })
}

// ─── System Tray ─────────────────────────────────────────────────────
function createTray() {
  const iconPath = path.join(__dirname, 'assets', 'icon.png')
  const icon = nativeImage.createFromPath(iconPath).resize({ width: 16, height: 16 })

  tray = new Tray(icon)
  tray.setToolTip('Mail Scraper')

  tray.on('click', () => {
    if (mainWindow) {
      mainWindow.show()
      mainWindow.focus()
    }
  })

  const contextMenu = Menu.buildFromTemplate([
    {
      label: 'Open Mail Scraper',
      click: () => {
        if (mainWindow) {
          mainWindow.show()
          mainWindow.focus()
        }
      },
    },
    { type: 'separator' },
    {
      label: 'Open in Browser',
      click: () => {
        shell.openExternal(FRONTEND_URL)
      },
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: () => {
        app.isQuiting = true
        if (backendProcess) backendProcess.kill()
        app.quit()
      },
    },
  ])

  tray.setContextMenu(contextMenu)
}

// ─── App Lifecycle ───────────────────────────────────────────────────
app.whenReady().then(async () => {
  createSplashWindow()
  await startBackend()
  try {
    await waitForBackend()
  } catch (err) {
    console.error('[Startup]', err.message, '— launching anyway')
  }
  createWindow()
  createTray()
  if (splashWindow && !splashWindow.isDestroyed()) splashWindow.close()
  mainWindow.show()
})

app.on('window-all-closed', (e) => {
  e.preventDefault() // keep running in tray
})

app.on('before-quit', () => {
  if (backendProcess) backendProcess.kill()
})

app.on('activate', () => {
  // macOS dock click
  if (mainWindow) mainWindow.show()
})
