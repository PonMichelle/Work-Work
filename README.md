# QS Workspace

A collaborative Quantity Surveying workspace (BOQ, rate build-ups / BUR, cost codes,
tender summary, and a sub-con quote approval workflow). Built with React + Vite.

## Live app

Once GitHub Pages is enabled (see below), the app is available at:

```
https://<your-github-username-or-org>.github.io/qs-workspace/
```

> **Data note:** the app stores data in each user's browser (`localStorage`).
> Data is **not** shared between different people or devices — each teammate
> has their own local copy. To make data truly shared/live across the team,
> the `window.storage` layer in `src/storage.js` would need to be pointed at a
> real backend (e.g. Firebase, Supabase, or a small API). Happy to add that next.

## Run locally

Requires [Node.js](https://nodejs.org) LTS (v18+).

```bash
npm install
npm run dev      # dev server at http://localhost:5173
npm run build    # production build into dist/
npm run preview  # preview the production build
```

## Deploying

This repo includes a GitHub Actions workflow (`.github/workflows/deploy.yml`)
that builds and deploys to GitHub Pages on every push to `main`.

To enable it: in GitHub, go to **Settings → Pages → Build and deployment →
Source = GitHub Actions**. (Pages on a private repo requires a paid GitHub plan.)

## Project structure

```
src/
  main.jsx      # entry point; installs the storage shim, mounts App
  storage.js    # localStorage-backed window.storage shim
  App.jsx       # the full application
```
