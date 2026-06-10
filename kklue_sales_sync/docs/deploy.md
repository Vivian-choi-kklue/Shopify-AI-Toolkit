# Deployment

This project is designed to run from a private GitHub repository.

## 1. Create Shopify Admin API Credentials

In Shopify Admin:

1. Go to Apps and sales channels.
2. Open Develop apps.
3. Create or open a custom app for automation.
4. Add Admin API scopes:
   - `read_orders`
   - `read_returns`
   - `read_customers`
   - `read_locations`
   - `read_inventory`
   - `read_users` if Shopify allows it
5. Install the app and copy the Admin API access token.

## 2. Add GitHub Secrets

In the private GitHub repo:

Settings > Secrets and variables > Actions > New repository secret

Add:

- `SHOPIFY_STORE_DOMAIN`
- `SHOPIFY_ADMIN_TOKEN`

## 3. Run Manually Once

Open Actions > Daily Shopify Sales Sync > Run workflow.

After it runs, check:

- `data/YYYY/YYYY MONTH Sales Data.xlsx`
- `data/YYYY/summary-YYYY-MM-DD.json`

If `manual_review` contains POS employee items, Shopify did not expose staff names to the API; those orders need note rules or manual review.
