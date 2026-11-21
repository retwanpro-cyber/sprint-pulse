# Sprint Pulse

Azure DevOps extension: sprint health, workload and risk insights for Azure Boards.

## Build the VSIX with GitHub Actions

1. Create a new GitHub repository and push the contents of this folder to the `main` branch.
2. Open **Actions → Build VSIX → Run workflow** (it also runs on every push).
3. When it finishes, download the **sprint-pulse-vsix** artifact. Unzip it to get the `.vsix` file.

## Before publishing

1. Create a publisher at https://marketplace.visualstudio.com/manage.
2. In `vss-extension.json`, replace `your-publisher-id` with your publisher ID.
3. Increase `version` for every new release.
4. Upload the `.vsix` in the Marketplace management portal.

## Marketplace tips

- Add 3–5 screenshots (1366×768 PNG) in `images/` and list them in `vss-extension.json` under `"screenshots"`:
  `"screenshots": [{ "path": "images/shot1.png" }]`
- Keep the name short and keyword-rich, and fill `tags` with terms users actually search for.
- Add `"repository"`, `"links"` (support, getstarted, learn) and `"badges"` for trust.

## Local build

```
npm install
npm run all
```
The VSIX is created in `out/`. The extension is browser-based, so it is OS-independent.
