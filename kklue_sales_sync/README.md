# KKLUE Shopify Sales Sync

Daily cloud workflow for rebuilding the current month's KKLUE sales workbook from Shopify Admin API data.

## Required GitHub Secrets

- `SHOPIFY_STORE_DOMAIN`: example `kklue-jewelry.myshopify.com`
- `SHOPIFY_ADMIN_TOKEN`: Shopify Admin API access token from a custom app

Recommended Shopify Admin scopes:

- `read_orders`
- `read_returns`
- `read_customers`
- `read_locations`
- `read_inventory`
- `read_users` if Shopify allows it, for POS staff names

If staff fields are not available, the workflow still runs and flags missing employee values for review.

## Output

The workflow writes:

- `data/YYYY/2026 JUNE Sales Data.xlsx`
- `data/YYYY/summary-YYYY-MM-DD.json`

The workflow runs daily at 09:30 Asia/Shanghai.
