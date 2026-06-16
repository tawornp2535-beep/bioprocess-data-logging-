# Deployment steps (Render + GitHub)

1. Create a GitHub repository and push this project:

```bash
git init
git add .
git commit -m "Initial commit for deployment"
git branch -M main
git remote add origin <your-github-repo-url>
git push -u origin main
```

2. On Render.com (recommended):
  - Create a new Web Service.
  - Connect your GitHub repository and select `main` branch.
  - Set the build command to: `npm ci && npm run build`
  - Set the start command to: `npm start` (or leave default since `start` exists).
  - Add an environment variable `MONGODB_URI` with your MongoDB Atlas connection string.
  - Deploy and Render will provide a public URL.

3. If you prefer Docker:
  - Push to a registry (Docker Hub / GitHub Container Registry) and deploy to any provider that supports containers.

4. Notes:
  - `server.js` serves the `dist/` build output. Ensure `npm run build` succeeds before starting.
  - If using MongoDB Atlas, allow Render's outbound IPs or use SRV connection string (normally works without IP allow-list).
