CSE 65-C PWA — GitHub Pages Ready

Upload/replace these files in the ROOT of your GitHub Pages repository:
- index.html
- manifest.webmanifest
- service-worker.js
- icons/ (whole folder)
- bus image.jpg / bus image-02.jpg / bus image-03.jpg / bus image-04.jpg

KEEP your existing files in the same repository:
- nayeem.jpg
- sabah.jpeg
- sabbir.jpeg
- amirul.jpeg

Then:
1. GitHub -> Settings -> Pages -> deploy the branch/root.
2. Open the HTTPS GitHub Pages URL in Android Chrome.
3. Wait a few seconds. Tap the green download icon in the portal header or Chrome menu (⋮) -> Install app.
4. The portal opens as a standalone app from the phone home screen.

Notes:
- Google Apps Script API requests are NOT cached by the service worker, so routine/student data stays live.
- If Chrome does not show Install immediately, hard refresh once and revisit the page.
- PWA install requires HTTPS; GitHub Pages provides HTTPS automatically.
