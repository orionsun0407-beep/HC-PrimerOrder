# Deploy HC PrimerOrder

## GitHub Pages

1. Create a public GitHub repository named `HC-PrimerOrder`.
2. Push this folder to the repository's `main` branch.
3. In GitHub, open `Settings -> Pages`.
4. Set `Build and deployment` to `Deploy from a branch`.
5. Choose `main` and `/ (root)`, then save.
6. The public URL will usually be:

   `https://orionsun0407-beep.github.io/HC-PrimerOrder/`

## Notes

- This is a static site. No server is needed.
- All primer/order files are processed in the visitor's browser.
- The ExcelJS browser bundle is vendored in `vendor/exceljs.min.js`, so the site does not depend on an external CDN at runtime.
