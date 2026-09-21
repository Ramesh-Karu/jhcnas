# Nextcloud WebDAV on TrueNAS SCALE Configuration Guide

## 1. Prerequisites on TrueNAS SCALE
1. Log into your TrueNAS SCALE WebUI (`https://truenas.local`).
2. Verify Nextcloud App is running under **Applications**.
3. Ensure the dataset where Excel files will reside has proper POSIX or NFSv4 permissions assigned to `www-data` (UID 33).

## 2. Generating a Dedicated Nextcloud App Password
> **Crucial Security Mandate:** Never use your personal Nextcloud user password. Always generate a dedicated App Password so it can be revoked independently without compromising your account.

1. In Nextcloud, click your User Profile Avatar (top right) -> **Personal Settings**.
2. Navigate to **Security** in the left navigation sidebar.
3. Scroll down to **Devices & sessions**.
4. Enter an App name: `Coolify-Excel-Sync`.
5. Click **Create new app password**.
6. Nextcloud will generate a 24-character token (e.g. `k8sL-99zA-pQ21-m5Nx-48vC`).
7. Copy this password into `NEXTCLOUD_APP_PASSWORD`.

## 3. WebDAV URL Format
In Nextcloud, WebDAV endpoints follow this exact pattern:
```
https://[YOUR_NEXTCLOUD_DOMAIN]/remote.php/dav/files/[USERNAME]/
```
Example:
```
https://cloud.example.com/remote.php/dav/files/excel-sync/
```

## 4. Source Folder Setup
1. In Nextcloud Files, create a dedicated folder named `ExcelImports` (or your chosen name).
2. Share access only with the `excel-sync` user.
3. Set `NEXTCLOUD_FOLDER=/ExcelImports`.

## 5. Connection Verification Matrix
The dashboard's **Test Nextcloud Connection** tool runs 4 verification checks:
1. **URL Reachable**: Resolves DNS and receives HTTP 200/207 response.
2. **Authentication**: Confirms HTTP Basic Auth credentials over HTTPS.
3. **Folder Accessible**: Executes WebDAV `PROPFIND` with `Depth: 1` on target folder.
4. **Excel Detection**: Verifies `.xlsx` and `.xlsm` files can be enumerated.
