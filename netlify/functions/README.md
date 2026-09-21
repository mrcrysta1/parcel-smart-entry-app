`api.mjs` is the shared multi-user API: it holds each uploaded workbook in
Netlify Blobs so several people can search, edit and add records in the same
file at once.

Netlify Blobs needs no account, connection string or environment variable -
credentials are injected into the function at runtime.

Run it locally exactly as it runs in production (Blobs emulated):

    npm install
    npx netlify dev
