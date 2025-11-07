# 🏗️ Building KRA POST PORTUM TOOL Desktop App

## ✅ Your App is Already a Desktop App!

This is a fully functional **Electron desktop application**. What you need is to **package it** into a distributable installer.

---

## 📦 **Build Instructions**

### **1. Install Dependencies (if not already done)**

```bash
npm install
```

### **2. Build the Desktop Installer**

#### **For Windows (.exe installer + portable):**

```bash
npm run build
```

This will create:

- `dist/KRA POST PORTUM TOOL Setup 1.0.0.exe` - **Installer**
- `dist/KRA POST PORTUM TOOL 1.0.0.exe` - **Portable** (no install needed)

#### **For Production Build:**

```bash
npm run dist
```

---

## 📁 **Output Files**

After building, check the `dist/` folder:

```
dist/
├── KRA POST PORTUM TOOL Setup 1.0.0.exe    # Windows Installer (NSIS)
├── KRA POST PORTUM TOOL 1.0.0.exe          # Windows Portable
└── win-unpacked/                            # Unpacked files (for testing)
```

---

## 🎯 **Installer Features**

### **NSIS Installer** (.exe)

- ✅ Custom installation directory
- ✅ Desktop shortcut
- ✅ Start Menu shortcut
- ✅ Uninstaller
- ✅ Admin privileges (for browser automation)

### **Portable Version**

- ✅ No installation required
- ✅ Run directly from USB/folder
- ✅ Self-contained

---

## 🚀 **Distribution**

### **Option 1: Direct Distribution**

1. Build the app: `npm run build`
2. Share the installer: `dist/KRA POST PORTUM TOOL Setup 1.0.0.exe`
3. Users double-click to install

### **Option 2: Portable Distribution**

1. Use the portable exe: `dist/KRA POST PORTUM TOOL 1.0.0.exe`
2. No installation needed
3. Users can run it from any folder

---

## 📋 **Pre-Build Checklist**

### **Required:**

- [X] ✅ Electron configured (main.js)
- [X] ✅ UI files (index-new.html, renderer-new.js, styles-new.css)
- [X] ✅ Automation scripts (automations/*.js)
- [X] ✅ electron-builder installed
- [ ] ⚠️ Icon files (assets/icon.ico for Windows)

### **Optional but Recommended:**

- [ ] App icon (256x256 PNG or ICO)
- [ ] Code signing certificate (for trusted installer)
- [ ] Version number in package.json
- [ ] Publisher name

---

## 🖼️ **App Icons**

### **Create Icons Folder:**

```bash
mkdir assets
```

### **Add Icon Files:**

- **Windows**: `assets/icon.ico` (256x256px)
- **macOS**: `assets/icon.icns`
- **Linux**: `assets/icon.png` (512x512px)

### **Generate ICO from PNG:**

If you have a PNG logo, convert it to ICO:

- Use online tool: https://convertio.co/png-ico/
- Or use ImageMagick: `convert logo.png -define icon:auto-resize=256,128,64,48,32,16 icon.ico`

---

## 🔧 **Build Scripts**

### **Available Commands:**

```bash
# Run in development mode
npm start

# Run with auto-reload and DevTools
npm run dev

# Build installer (Windows)
npm run build

# Build for distribution (all platforms)
npm run dist
```

---

## 💻 **System Requirements**

### **Development:**

- Node.js 16+
- npm 7+
- Windows 10+ / macOS 10.13+ / Linux

### **Runtime (Built App):**

- Windows 10/11 (64-bit)
- Chrome/Chromium browser installed (for Playwright)
- 4GB RAM minimum
- 500MB disk space

---

## 🎨 **Customization**

### **Change App Name:**

Edit `package.json`:

```json
{
  "name": "your-app-name",
  "productName": "Your App Display Name",
  "version": "1.0.0"
}
```

### **Change Window Size:**

Edit `main.js`:

```javascript
mainWindow = new BrowserWindow({
    width: 1600,  // Change width
    height: 1000, // Change height
    ...
});
```

### **Change App ID:**

Edit `package.json` → `build.appId`:

```json
{
  "build": {
    "appId": "com.yourcompany.your-app"
  }
}
```

---

## 🐛 **Troubleshooting**

### **Build Fails:**

```bash
# Clear cache and rebuild
rm -rf dist/
rm -rf node_modules/
npm install
npm run build
```

### **"Icon not found" Error:**

Create a placeholder icon or remove icon reference:

```json
"win": {
  // Remove or comment out:
  // "icon": "assets/icon.ico"
}
```

### **"Permission Denied" Error:**

Run as administrator:

```bash
# Windows PowerShell (as Admin)
npm run build

# Or use portable build
npm run dist -- --win portable
```

---

## 📦 **File Size**

Expect the installer to be **~300-500MB** due to:

- Electron runtime (~150MB)
- Chromium browser (~100MB)
- Node modules (~150MB)
- Playwright browsers (~100MB)

This is normal for Electron apps with browser automation!

---

## 🎯 **Quick Start Guide for Users**

### **Installation:**

1. Download `KRA POST PORTUM TOOL Setup 1.0.0.exe`
2. Double-click to run installer
3. Follow installation wizard
4. Launch from Desktop or Start Menu

### **First Run:**

1. Enter KRA PIN and Password
2. Click "Validate Credentials"
3. Select output folder
4. Start extracting data!

---

## 📝 **Version Updates**

To release a new version:

1. Update version in `package.json`:

```json
{
  "version": "1.1.0"
}
```

2. Rebuild:

```bash
npm run build
```

3. New installer will be: `KRA POST PORTUM TOOL Setup 1.1.0.exe`

---

## 🔐 **Code Signing (Optional)**

For a trusted installer without Windows SmartScreen warnings:

1. Get a code signing certificate
2. Add to package.json:

```json
{
  "win": {
    "certificateFile": "path/to/certificate.pfx",
    "certificatePassword": "your-password"
  }
}
```

---

## ✅ **Build Checklist**

- [ ] All dependencies installed (`npm install`)
- [ ] App tested in dev mode (`npm start`)
- [ ] Version updated in package.json
- [ ] Icon file created (optional)
- [ ] Build completed successfully (`npm run build`)
- [ ] Installer tested on clean Windows machine
- [ ] Portable version tested
- [ ] README created for users
- [ ] Distribution method decided

---

## 📧 **Support**

For build issues:

1. Check Node.js version: `node --version` (should be 16+)
2. Check npm version: `npm --version` (should be 7+)
3. Clear cache: `npm cache clean --force`
4. Reinstall: `rm -rf node_modules && npm install`

---

**Your app is ready to build! Just run `npm run build` 🚀**
