# Avalie Team Icon Vault

A PowerPoint Office 365 task pane add-in for searching and inserting SVG icons from 7 icon libraries directly onto slides.

## Icon Libraries (~40,000 icons)

- **Fluent UI** — Microsoft's design system icons
- **Tabler** — 5,400+ pixel-perfect icons
- **Remix** — open-source icon set
- **FA6 Solid / Regular** — Font Awesome 6
- **Lucide** — community-driven fork of Feather Icons
- **Heroicons** — by the Tailwind CSS team

## Setup

```bash
npm install
npm run build
npm start        # starts HTTPS dev server on localhost:3000
```

## Sideloading into PowerPoint

### Mac Desktop
Copy `manifest.xml` to:
```
~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```
Then restart PowerPoint → Insert → My Add-ins.

### Windows Desktop
Copy `manifest.xml` to:
```
%USERPROFILE%\AppData\Local\Microsoft\Office\16.0\Wef\
```
Then restart PowerPoint → Insert → My Add-ins.

### Office on the Web
Insert → Office Add-ins → Upload My Add-in → select `manifest.xml`.

## GitHub Pages Deployment

1. Push this repo to GitHub
2. Go to **Settings → Pages → Source** → select `main` branch, `/dist` folder
3. Update all URLs in `manifest.xml` from `https://localhost:3000` to your Pages URL (e.g. `https://yourusername.github.io/repo-name`)
4. Distribute the updated `manifest.xml` to your team

## Usage

1. Click **Icon Vault** in the **Avalie Icons** ribbon tab
2. Search by keyword (e.g. `arrow`, `chart`, `user`)
3. **Click** an icon → inserts as editable SVG on the current slide
4. **Shift+click** → copies SVG to clipboard
5. Toggle pack filters to narrow results
