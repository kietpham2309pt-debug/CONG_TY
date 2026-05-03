# Wellhome Order Sync

Pipeline tự động Haravan → GHN → Fulfill cho shop Wellhome (wellhome.asia), tích hợp FFM Onflow.

## Cấu trúc

```
wellhome-order-sync/
├── apps-script/           # Code Apps Script (project "Wellhome Order Sync" trên admin@khomes.com.vn)
│   ├── phase1.js                  # Pull đơn Haravan → tab "tracking haravan"
│   ├── phase2_ghn.js              # Tạo đơn GHN + doPost handler (entry point Web App)
│   ├── phase2_5_failed_summary.js # Daily 8h: email 4 nhóm đơn cần xử lý
│   ├── phase3_status.js           # Update GHN status 30p
│   ├── phase4_fulfill.js          # Fulfill Haravan khi đơn giao xong
│   ├── phase5_print.js            # In tem GHN
│   ├── phase6_inventory.js        # Pull tồn kho Haravan daily 6h
│   ├── phase7_daily_recon.js      # Daily recon 18h + KL quy đổi alert
│   ├── phase8_auto_retry.js       # Auto retry đơn lỗi GHN daily 9h
│   ├── phase9_weekly_report.js    # Weekly report Mon 9h
│   ├── phase11_health_monitor.js  # Health check Sun 23h
│   ├── phase12_sync_scheme.js     # Sync Scheme master daily 5h
│   ├── phase15_monthly_report.js  # Monthly report ngày 1
│   ├── phase16a_push_oms.js       # Push đơn → OMS Onflow (config qua ScriptProperties)
│   ├── phase16b_oms_webhook.js    # OMS webhook receiver + public lookup HTML
│   ├── phase17_cod_recon.js       # Đối soát COD GHN daily 19h
│   ├── phase18_fee_alert.js       # Cảnh báo phí ship cao daily 19:15
│   ├── phase19_pickup.js          # Pickup workflow (notify 9h + fulfill 1h)
│   ├── phase20_kol_perf.js        # KOL aggregation daily 19:30
│   ├── phase21_recover_order.js   # Recover đơn xóa nhầm
│   ├── _setup_runner.js           # runIncrementalSetup + setupAllTriggers
│   ├── Chưa có tên.js             # Bound script template (paste vào Sheet)
│   └── appsscript.json            # Manifest
└── cloudflare-worker/     # Proxy CORS cho OMS webhook
    ├── worker.js
    └── wrangler.toml
```

## Pipeline

```
Haravan (wellhome.asia) ─Phase 1 1h─→ Sheet "tracking haravan"
                                          ├─ Phase 2 1h ──→ GHN (tạo vận đơn)
                                          ├─ Phase 16a 30p → OMS Onflow (FFM soạn hàng)
                                          ▼
                                      FFM xử lý
                                          │
                                          ├─ Phase 16b ←── webhook OMS Onflow (qua Cloudflare Worker)
                                          ▼
                                      Phase 3 30p: poll GHN status
                                      Phase 4 30p: fulfill Haravan khi giao xong
```

## Setup lần đầu (cho người mới clone)

### 1. Apps Script (clasp)
```bash
npm install -g @google/clasp
clasp login
cd apps-script
echo '{"scriptId": "YOUR_SCRIPT_ID"}' > .clasp.json
clasp push --force
```

### 2. Cấu hình token (qua Web App POST)
```powershell
# Set Haravan token
$body = @{ secret = 'YOUR_WEBAPP_SECRET'; action = 'set_haravan_token'; token = '...' }
Invoke-RestMethod -Method Post -Uri '<WEB_APP_URL>' -Body $body

# Set GHN token, OMS API key tương tự...
```

Chi tiết các action trong `phase2_ghn.js doPost`.

### 3. Cloudflare Worker
```bash
cd cloudflare-worker
npm install -g wrangler
wrangler login
wrangler secret put APPS_SCRIPT_URL  # paste URL Web App
wrangler secret put OMS_TOKEN         # paste token webhook
wrangler deploy
```

## Triggers active (16 cron)

| Cron | Hàm | Phase |
|---|---|---|
| 1h | syncOrders | 1 |
| 1h | createGhnOrders | 2 |
| 30p | updateGhnStatuses | 3 |
| 30p | fulfillHaravanOrders | 4 |
| 30p | pushOrdersToOms | 16a |
| 1h | fulfillPickupOrders | 19 |
| daily 5h | syncSchemeToLocal | 12 |
| daily 6h | syncHaravanInventory | 6 |
| daily 8h | sendDailyGhnFailedSummary | 2.5 |
| daily 9h | autoRetryFailedGhnOrders + sendMonthlyReport (date=1) + notifyNewPickupOrders | 8/15/19 |
| daily 18h | runDailyRecon | 7 |
| daily 19h | runCodRecon | 17 |
| daily 19:15 | runFeeAlertReport | 18 |
| daily 19:30 | aggregateKolPerformance | 20 |
| Mon 9h | sendWeeklyReport | 9 |
| Sun 23h | runHealthMonitor | 11 |

## Bảo mật

- KHÔNG commit token/secret. Tất cả lưu trong Apps Script ScriptProperties hoặc Cloudflare Worker secrets.
- Secret hardcode trong code chỉ là placeholder `REPLACE_WITH_*` — phải replace trước khi run setup.
- Repo PUBLIC nên KHÔNG push file `.clasp.json` (chứa scriptId).

## License
Internal use — K-Homes / Wellhome.
