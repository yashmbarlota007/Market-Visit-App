# 📍 Yash Marketing - Market Visit OS

Live production market visit tracking, dealer location geofence locking, and anti-proxy verification portal.

---

## ⚡ Core Operational Features

1. **📡 Real-Time Staff Live Radar (Owner Command Center):**
   - Live interactive radar map (OpenStreetMap / Leaflet.js).
   - Shows active staff field location, speed (km/h), distance from HQ, and time elapsed since last ping.
   - Direct "Open in Google Maps" link for every staff member.

2. **🔒 Dealer GPS Location Lock:**
   - Lock permanent store GPS coordinates (Latitude & Longitude) for every dealer.
   - Can be locked directly by staff on first visit (with store photo proof) or by the Owner.
   - Once locked, the store location is permanently fixed to prevent spoofing.

3. **🛡️ Strict 50-Meter Geofenced Check-In (Anti-Proxy Punch):**
   - High-accuracy Haversine distance engine calculates exact radial distance between staff and dealer.
   - **If distance ≤ 50m**: Check-in is approved! Camera opens for Store Front photo capture, discussion notes, and order/payment logging.
   - **If distance > 50m**: **STRICTLY REJECTED**. Visual alert shows exact distance offset (e.g. *140m away*) and logs breach to the `ProxyAuditLog` sheet.

---

## 🚀 Instant Deployment with Vercel

1. Log in to [Vercel](https://vercel.com).
2. Click **Add New...** -> **Project**.
3. Import **`yashmbarlota007/Market-Visit-App`**.
4. Framework Preset: **Other** (Root directory: `./`).
5. Click **Deploy**.

---

## ⚙️ Google Apps Script (GAS) Sync

- **Script ID:** `1PJzNM8iF9vMCEQMXlq8A8eQ-CQkUwVHllAq9ydaLySFBtJ-e5D44wubw`
- Sync live code from this GitHub repository to your Google Apps Script project directly online using the **Google Apps Script GitHub Assistant** Chrome extension.
