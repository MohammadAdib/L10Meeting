# L10 Meeting Manager

A web app for running EOS L10 meetings. Runs entirely in the browser with no server required — data is stored directly on your computer using the File System Access API.

## Features

- Department management with people lists
- L10 meeting workflow (segue, scorecard, OKR review, headlines, to-do review, IDS, conclude)
- Scorecard and OKR tracking across meetings
- Auto-save to Excel files on your local filesystem
- Custom company logo support
- Works with any cloud-synced folder (OneDrive, Google Drive, Dropbox)

## Requirements

- A Chromium-based browser (Chrome, Edge, Opera)

## Development

```
npm install
npm run dev
```

Open `http://localhost:5173/CompanyTools/`.

## Production Build

```
npm run build
```

Deploy the `dist/` folder to any static host (GitHub Pages, Vercel, Netlify, etc.).

## How It Works

On first visit, the app asks you to select a folder on your computer. All meeting data (Excel files, people lists, logo) is stored directly in that folder. The folder choice is remembered in the browser via IndexedDB so you only need to pick it once.
