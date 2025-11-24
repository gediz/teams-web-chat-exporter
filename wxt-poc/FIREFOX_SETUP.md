# Firefox Add-on Store - Setup Checklist

Quick checklist to fix Firefox Add-ons store warnings.

---

## ✅ Step 1: Set Your Add-on ID

**Edit `wxt.config.ts` line 18:**

```typescript
id: 'teams-chat-exporter@yourdomain.com',
```

**Change to one of these formats:**

✅ **Your email:**
```typescript
id: 'teams-chat-exporter@yourname.com',
```

✅ **Your domain:**
```typescript
id: 'teams-chat-exporter@github.io',
```

✅ **Generic (if no domain/email):**
```typescript
id: '{12345678-1234-1234-1234-123456789012}',
```
Use a UUID generator: https://www.uuidgenerator.net/

**Important**:
- Must be unique across all Firefox add-ons
- Cannot change once published
- Format: `name@domain` or `{uuid}`

---

## ✅ Step 2: Verify Data Collection Settings

Already configured in `wxt.config.ts`:

```typescript
data_collection_permissions: {
  required: ["none"], // ✅ Correct for this extension
},
```

**What this means:**
- This extension does NOT collect user data
- All exports save locally to user's device
- No analytics, no telemetry, no tracking
- `["none"]` = no data collection required

---

## ✅ Step 3: Rebuild and Re-upload

```bash
# Rebuild Firefox version
npm run zip:firefox

# Upload new ZIP to Firefox Add-ons
# Go to: https://addons.mozilla.org/developers/addon/submit/upload-listed
```

---

## Warnings Resolved

✅ **Warning 1: Missing add-on ID**
```
The "/browser_specific_settings/gecko/id" property should be specified
```
**Fixed**: Added `id: 'teams-chat-exporter@yourdomain.com'`
**Action**: Replace `yourdomain.com` with your actual domain/email

✅ **Warning 2: Missing data collection permissions**
```
The "/browser_specific_settings/gecko/data_collection_permissions" property is required
```
**Fixed**: Added `data_collection_permissions: { collect_data: false }`
**Action**: No change needed (extension doesn't collect data)

---

## Current Config

Your `wxt.config.ts` now has:

```typescript
browser_specific_settings: {
  gecko: {
    id: 'teams-chat-exporter@yourdomain.com', // ⚠️ CHANGE THIS
    strict_min_version: '109.0',
    data_collection_permissions: {
      required: ["none"], // ✅ Correct
    },
  },
},
```

---

## Next Steps

1. **Update ID** in `wxt.config.ts` (line 18)
2. **Rebuild**:
   ```bash
   npm run zip:firefox
   ```
3. **Upload** to Firefox Add-ons store
4. **Warnings should be gone!** ✅

---

## Choosing an Add-on ID

### Option 1: Use Your Email (Recommended)
```typescript
id: 'teams-chat-exporter@yourname.com',
```
✅ Easy to remember
✅ Shows you own it

### Option 2: Use Your Domain
```typescript
id: 'teams-chat-exporter@yourdomain.com',
```
✅ Professional
✅ Good if you have a website

### Option 3: Use UUID (If no email/domain)
```typescript
id: '{a1b2c3d4-e5f6-7890-abcd-ef1234567890}',
```
✅ Guaranteed unique
❌ Hard to remember

**Generate UUID**: https://www.uuidgenerator.net/version4

---

## Important Notes

⚠️ **Cannot change ID after publishing!**
- Choose carefully
- Must be unique across ALL Firefox add-ons
- Used for updates and user installs

✅ **Data collection setting is correct**
- Your extension saves everything locally
- No data sent to servers
- `collect_data: false` is accurate

---

## Verification

After rebuilding, check the generated manifest:

```bash
# Check generated manifest
cat .output/firefox-mv2/manifest.json | grep -A 5 "browser_specific_settings"
```

Should show:
```json
"browser_specific_settings": {
  "gecko": {
    "id": "teams-chat-exporter@yourname.com",
    "strict_min_version": "109.0",
    "data_collection_permissions": {
      "required": ["none"]
    }
  }
}
```

---

**Ready to upload?** Follow [DEPLOYMENT_GUIDE.md - Firefox Add-ons](DEPLOYMENT_GUIDE.md#firefox-add-ons-amo)
