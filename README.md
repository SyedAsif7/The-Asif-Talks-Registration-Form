# The Asif Talks — Episode #1 Registration Page

A dedicated, high-converting, mobile-responsive live audience registration page for **The Asif Talks (Episode #1)** featuring **Shri Sanjaysinh Chavan (IAS)**, District Collector & DM Parbhani.

---

## 📁 Project Structure

```
.
├── index.html              # Streamlined registration landing page
├── assets/                 # High-resolution logos & speaker images
│   ├── asif_talks_host.jpg # Host & Podcast branding graphic
│   ├── collector_chavan.png# Shri Sanjaysinh Chavan (IAS) portrait
│   ├── dcode_logo.jpg      # DCode Developers Club logo
│   ├── ssiems_logo.png     # SSIEMS Parbhani emblem
│   └── vertex_logo.jpg     # Vertex Institute of Technology logo
├── google-apps-script.js   # Production-ready Apps Script backend with LockService
└── README.md               # Quickstart and deployment instructions
```

---

## ⚡ Quickstart Guide (3 Steps)

### Step 1: Set Up Google Sheet Backend
1. Open [Google Sheets](https://sheets.new) and create a new sheet.
2. In **Row 1**, enter these 12 column headers:
   `Timestamp` | `Pass ID` | `Student Name` | `Mobile Number` | `Email Address` | `Gender` | `College / Institute / Organization` | `City / Location` | `Expectations` | `How Heard` | `Question for Collector` | `Photo/Video Consent`
3. Click **Extensions $\rightarrow$ Apps Script**.
4. Paste all code from [`google-apps-script.js`](file:///c:/Users/as/Desktop/Reg/google-apps-script.js) and save (Ctrl+S).
5. Click **Deploy $\rightarrow$ New deployment**:
   - **Select type:** Web app
   - **Execute as:** `Me`
   - **Who has access:** `Anyone`
6. Click **Deploy**, grant permissions, and copy the generated **Web App URL**.

---

### Step 2: Connect Frontend
1. Open [`index.html`](file:///c:/Users/as/Desktop/Reg/index.html).
2. Locate line **399**:
   ```javascript
   const scriptURL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
3. Replace with your copied Web App URL.

---

### Step 3: Deploy Online (Free)
- **Vercel:** Import your GitHub repo or drag-and-drop the folder onto [Vercel](https://vercel.com).
- **GitHub Pages:** Go to repository **Settings $\rightarrow$ Pages $\rightarrow$ Deploy from `main` branch**.

---

## 📝 Form Fields Included
1. **Student Name \*** (`name` - Enter your full name)
2. **Mobile Number \*** (`phone` - 10-digit mobile number)
3. **Email Address \*** (`email` - abc@gmail.com)
4. **Gender \*** (`gender` - Male, Female, Other, Prefer not to say)
5. **College / Institute / Organization Name \*** (`college`)
6. **City / Location \*** (`city`)
7. **What are you expecting from The ASIF Talks? (Optional)** (`expectations` - Leadership insights, Career guidance, New perspectives, Networking, Knowledge about governance, Youth empowerment, Other)
8. **How did you hear about this event? \*** (`source` - WhatsApp, Instagram, College/School, Friend, Event Team, Other)
9. **Your Question for District Collector Shri Sanjaysinh Chavan (IAS) (Optional)** (`question`)
10. **Photo/Video Consent \*** (`consent` - Checkbox agreement)
