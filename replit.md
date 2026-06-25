# Bank Statement Analyzer

A professional Excel Add-in and standalone web application for analyzing, categorizing, and visualizing financial data from bank statements.

## Overview

- **Frontend**: React 18 + TypeScript + Vite, served on port 5000
- **Backend**: Node.js + Express API, served on port 3001
- **Vite proxies** `/api` requests to the backend automatically in development

## Features

- Upload CSV, XLSX, or PDF bank statements
- Automatic transaction categorization and financial health scoring
- Charts and visualizations via Recharts
- Email report delivery via SMTP (Nodemailer)
- Pro subscription system verified via USDT blockchain payments (Tron, BSC, Ethereum)
- Admin dashboard for license management, revenue stats, and settings
- Microsoft Excel Add-in support via Office.js

## Running the Project

The `Start application` workflow runs both backend and frontend:
```bash
npm run dev
```

## Environment Variables / Secrets

Set these in the Replit Secrets panel:

| Variable | Required | Description |
|---|---|---|
| `ADMIN_PASSWORD` | Yes | Password for the admin dashboard at `/admin` |
| `USDT_WALLET_ADDRESS` | For payments | Your USDT wallet address to receive payments |
| `USDT_NETWORK` | For payments | `tron`, `bsc`, or `eth` (default: `tron`) |
| `TRONGRID_API_KEY` | Optional | TronGrid API key for Tron transaction verification |
| `BSCSCAN_API_KEY` | For BSC payments | BscScan API key |
| `ETHERSCAN_API_KEY` | For ETH payments | Etherscan API key |

SMTP settings for email delivery are configured via the Admin UI (Settings > Notifications).

## Data Storage

The server uses flat JSON files in `server/data/` for persistence:
- `licenses.json` — issued license keys
- `settings.json` — app appearance, payment config, plans, notifications
- `tickets.json` — support tickets

## User Preferences

- Use `concurrently` to run backend + frontend together in dev mode
