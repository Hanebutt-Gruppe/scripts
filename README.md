# Static Scripts

Pure static scripts for Windows and Linux.

---

## 🔧 Quick Execution

### 🐧 Linux
```bash
curl -fsSL https://raw.githubusercontent.com/hanebutt-gruppe/scripts/main/linux/<script>.sh | bash -s -- [args]
```

---

### 🪟 Windows

#### Option 1 — Execute from URL (inline)
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr https://raw.githubusercontent.com/hanebutt-gruppe/scripts/main/windows/<Script>.ps1 -UseBasicParsing | iex"
```

#### Option 2 — Download then run
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "iwr https://raw.githubusercontent.com/hanebutt-gruppe/scripts/main/windows/Fix-Dns.ps1 -OutFile Fix-Dns.ps1; powershell -NoProfile -ExecutionPolicy Bypass -File .\Fix-Dns.ps1"
```