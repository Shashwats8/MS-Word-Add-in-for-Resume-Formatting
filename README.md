# MS-Word-Add-in-for-Resume-Formatting

Resume formatting is widely used in the US staffing Industry and HR Tech. This add-in helps recruiters optimize their workflow by automatically formatting Word documents (resume style) using a set of instructions.

## What it does

- Applies Calibri 11 throughout the document
- Sets 1.15 line spacing and removes extra blank lines
- Converts list-like lines to bullet points
- Detects heading section titles (Experience, Education, Skills, Summary, etc.) and sets `Heading 2`
- Aligns date-containing paragraphs right
- Provides a button to format the whole doc or selected text

## Files in this repo

- `manifest.xml` - Office Add-in manifest
- `taskpane.html` - Add-in pane UI
- `taskpane.js` - JavaScript formatting logic
- `taskpane.css` - UI styles
- `package.json` - local dev tools and start scripts
- `.gitignore`

## Setup and run

1. Install dependencies

   ```bash
   npm install
   ```

2. Trust dev cert and run local HTTPS server

   ```bash
   npm run start
   ```

3. Sideload the add-in in Word (Windows or Mac):
   - Open Word, go to **Insert > Add-ins > My Add-ins > Manage My Add-ins > Upload My Add-in**
   - Choose `manifest.xml`

4. Open any doc and use the task pane buttons.

## Manual test instructions

1. Paste a sample resume text with section titles and dates.
2. Click **Format Document**.
3. Verify heading sections are styled as `Heading 2` and regular text is Normal.
4. Verify bullets are applied and spacing is normalized.

## Advanced improvements

- Add per-section template presets (US/UK/International resume)
- Expose custom style mapping UI
- Validate sections and generate a resume quality score
- Export to PDF once formatted

