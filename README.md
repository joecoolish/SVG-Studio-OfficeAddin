# IconCloud PowerPoint Add-in

A Microsoft **PowerPoint** Office Add-in (task pane) that connects to the
[IconCloud](https://iconcloud.design/) SVG collection API to **list, search, and
insert** SVG icons directly into the active presentation as native, scalable
vector graphics.

Built with **TypeScript + React + Fluent UI v9** and the standard Office Add-in
webpack tooling.

## Features

- 🔍 **Search / browse** icons from the IconCloud API.
- 🖼️ **Grid preview** of results (uses inline SVG or preview URL).
- ➕ **One-click insert** into the current slide via
  `Office.CoercionType.XmlSvg` (true vector, not rasterized).
- 🔐 **Pluggable auth** — Office SSO (Entra ID) with a manual bearer-token
  fallback for API exploration.

## Important: the API is Entra ID–protected

Every route on `iconcloud.design` (including `/swagger`, `/openapi.json`, etc.)
redirects to a **Microsoft Entra ID** sign-in page. The exact API surface is not
public, so the client is written to be **configurable** and **defensive**:

- Adjust base URL, endpoint paths, and response shape in
  [`src/services/config.ts`](src/services/config.ts).
- The response parser (`src/services/iconCloudApi.ts`) tolerates several common
  JSON conventions (`items` / `results` / `data` / `value`, `id` / `iconId`,
  etc.). Tighten it once the real responses are confirmed.

### Configuring endpoints

Edit `src/services/config.ts`:

```ts
export const config: IconCloudConfig = {
  baseUrl: "https://iconcloud.design",
  endpoints: {
    search: "/api/icons?search={query}&page={page}&pageSize={pageSize}",
    list: "/api/icons?page={page}&pageSize={pageSize}",
    svg: "/api/icons/{id}/svg",
  },
  svgDelivery: "text", // or "field" if the SVG markup is embedded in the icon JSON
  svgField: "svg",
  pageSize: 30,
};
```

Supported template tokens: `{query}`, `{page}`, `{pageSize}`, `{id}`.

### Authentication

Two strategies, resolved in [`src/services/auth.ts`](src/services/auth.ts):

1. **Office SSO (recommended for production)** — add a `<WebApplicationInfo>`
   section to `manifest.xml` with your Entra app registration ID and scopes.
   `Office.auth.getAccessToken()` is then used automatically.
2. **Manual token (exploration)** — when a request returns 401/403, the task
   pane shows a panel to paste a bearer token (e.g. captured from the existing
   web app's network calls).

## Getting started

```powershell
npm install

# Validate the manifest
npm run validate

# Start the dev server + sideload into PowerPoint desktop
npm start

# Or just run the dev server
npm run dev-server
```

`npm start` launches PowerPoint, trusts the dev certificate, and sideloads the
add-in. Open the **Home** tab → **Browse Icons** to show the task pane.

To stop debugging:

```powershell
npm stop
```

## Project structure

```
manifest.xml                     PowerPoint task-pane manifest (Host: Presentation)
src/
  services/
    config.ts                    Base URL, endpoint templates, response options
    auth.ts                      Token provider (Office SSO + manual token)
    iconCloudApi.ts              list() / search() / getSvg() + normalization
  office/
    insertSvg.ts                 Insert SVG into slide via XmlSvg coercion
  taskpane/
    index.tsx                    React entry (FluentProvider)
    taskpane.html
    components/                  App, Header, SearchBar, IconGrid, IconCard, TokenPanel
  commands/
    commands.ts / commands.html  Ribbon function-file target
assets/                          Add-in icons
```

## Deploying

1. Host the built `dist/` output on HTTPS.
2. Update `urlProd` in `webpack.config.js` and the URLs in `manifest.xml`.
3. Run `npm run build` and deploy `dist/`.

## Notes

- SVG insertion uses `Office.context.document.setSelectedDataAsync` with
  `Office.CoercionType.XmlSvg`, which PowerPoint renders as an editable vector
  shape.
- `AppDomains` in the manifest includes `https://iconcloud.design` so auth
  redirects stay within the add-in.
# SVG-Studio-OfficeAddin
