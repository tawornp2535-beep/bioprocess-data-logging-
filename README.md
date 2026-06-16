# React + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is not enabled on this template because of its impact on dev & build performances. To add it, see [this documentation](https://react.dev/learn/react-compiler/installation).

## Expanding the ESLint configuration

If you are developing a production application, we recommend using TypeScript with type-aware lint rules enabled. Check out the [TS template](https://github.com/vitejs/vite/tree/main/packages/create-vite/template-react-ts) for information on how to integrate TypeScript and [`typescript-eslint`](https://typescript-eslint.io) in your project.

## Friendly local URL (custom hostname)

If you'd like to open the app using a friendly name like `http://bioprocess.local:5173` instead of `http://localhost:5173`, follow these steps (requires administrator rights on Windows):

1. Open Notepad as Administrator and edit the hosts file: `C:\Windows\System32\drivers\etc\hosts`.
2. Add a line mapping the name to localhost (example):

```
127.0.0.1 bioprocess.local
```

3. Save the file. Windows may require elevated privileges to save changes.
4. Start the dev server as usual:

```powershell
npm install
npm run dev
```

5. Open the friendly URL in your browser:

```
http://bioprocess.local:5173/
```

Helper files added to the repository:
- `open_bioprocess.bat` — opens your default browser to `http://bioprocess.local:5173` (use after editing hosts file).
- `add-hosts-instructions.txt` — step-by-step instructions in Thai for editing `hosts` safely.

Note: Using `.local` may rely on mDNS on some networks; mapping via the hosts file above is the most reliable locally. If you want a true DNS name for other devices on the network, you need to configure your router or a DNS server.

