# Testing Guide

This app is a Tauri desktop frontend that relies on your configured Google Apps Script backend. The steps below show how to set up the environment, run the UI, and manually verify the tenant directory features that were recently added.

## Prerequisites
- Node.js 18+ and npm
- Rust toolchain with `cargo`
- Tauri CLI (`npm install -g @tauri-apps/cli`), or rely on `npx tauri` when using `npm run tauri`
- Access to the expected Apps Script endpoints (the `ext_config.json` file should already point the frontend to your deployment)

## Install dependencies
Run the install once after cloning or pulling changes:

```bash
npm install
```

## Start the app in development mode
Launch the Tauri dev server and desktop shell:

```bash
npm run tauri dev
```

The command opens the desktop window with hot reload. If the backend is unreachable, check your network/VPN and Apps Script deployment URL.

## Manual QA: View Tenants directory
Use these steps to exercise the new navigation item and modal editing flow.

1. Start the app, then click **View Tenants** in the top navigation.
2. Confirm the table loads tenant rows with Name and Wing/Floor columns.
3. Use the **Status** filter to toggle between **Active**, **Inactive**, and **All** and verify the table updates accordingly.
4. Use the search box to filter by tenant name; clearing the box should reset the rows.
5. Click **Refresh** to force a reload from the backend and ensure the table repopulates.
6. Click **Edit** on a row to open the modal.
   - Verify tenant fields populate (name, rent, wing/floor, status, contact details).
   - Switch to the **Family** tab and confirm related members show up (name, relation, age/gender if provided).
7. Update a tenant field (e.g., status or contact) and save; expect a success confirmation and the row to reflect the change after closing the modal.
8. Update or add a family member entry and save; confirm the modal shows the change and the backend accepts the update.
9. Close the modal with the **X** button and by clicking outside the modal to ensure both dismissal methods work.
10. Repeat with both active and inactive tenants to ensure state handling is consistent.

## Troubleshooting
- If `npm run tauri dev` fails because of missing Rust, install Rust via [rustup](https://rustup.rs/).
- If the UI shows empty data, verify your Apps Script deployment is reachable and returns tenant records.
- To clear cached builds, remove the `src-tauri/target` directory and rerun the dev command.
