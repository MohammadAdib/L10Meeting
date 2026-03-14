# L10 Meeting Tool

Internal meeting management tool for running EOS L10 meetings. Built for Titan Dynamics.

Features:
- Department management with people lists
- L10 meeting workflow (segue, scorecard review, OKR review, headlines, to-do review, IDS, conclude)
- Scorecard and OKR tracking across meetings
- Auto-save with debounce
- Excel export
- Standalone Windows executable

## Running Locally

```
npm install
npm run start
```

This builds the frontend and server, then starts it at `http://localhost:3847`.

### Development with Hot Reload

For frontend hot reload during development, run two terminals:

Terminal 1 — build and start the API server:
```
npm run build:server
node server-dist/index.cjs
```

Terminal 2 — start Vite dev server:
```
npm run dev
```

Then open `http://localhost:5173/CompanyTools/`. Frontend changes hot reload automatically. API calls are proxied to the server on port 3847.

## Building the Executable

```
npm run package
```

Or use the batch script:

```
build.bat
```

The executable will be output to `release/L10Meeting.exe`. Run it to start the server and auto-open the browser. Meeting data is stored in a `data/` folder next to the executable.
