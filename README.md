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

```bash
npm install
npm run build
node server-dist/index.cjs
```

This builds the frontend, compiles the server, and starts it at `http://localhost:3847`.

For development with hot reload (frontend only — needs the server running separately):

```bash
npm run dev
```

## Building the Executable

```bash
npm run package
```

Or use the batch script:

```bash
build.bat
```

The executable will be output to `release/L10Meeting.exe`. Run it to start the server and auto-open the browser. Meeting data is stored in a `data/` folder next to the executable.
