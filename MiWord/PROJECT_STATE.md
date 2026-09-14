name: Build Windows (TETORD)

on:
  push:
    branches: [ main, master ]
    tags: [ 'v*' ]
  pull_request:
    branches: [ main, master ]
  workflow_dispatch: {}

jobs:
  build-windows:
    runs-on: windows-latest

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Set up Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'

      - name: Install dependencies
        run: npm install

      - name: Build Windows installer (NSIS) and portable
        run: npm run build:release

      - name: Locate build output
        shell: pwsh
        run: Get-ChildItem -Path dist -Recurse | Select-Object FullName

      - name: Upload installer artifact
        uses: actions/upload-artifact@v4
        with:
          name: TETORD-Setup-x64
          path: dist/TETORD-Setup-*-x64.exe
          if-no-files-found: error

      - name: Upload portable artifact
        uses: actions/upload-artifact@v4
        with:
          name: TETORD-Portable-x64
          path: |
            dist/TETORD-*-x64.exe
            !dist/TETORD-Setup-*-x64.exe
          if-no-files-found: error
