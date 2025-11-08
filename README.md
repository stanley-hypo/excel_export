# Excel Merge (Node.js library + page)

Replace `{{placeholders}}` in an Excel `.xlsx` template with JSON data (mail-merge style).

## Features
- Read an Excel template and a JSON object
- Replace placeholders like `{{name}}`, `{{order.total}}`, `{{items[0].price}}`
- Simple web page to paste JSON, upload template, and download result

## Quick Start

1) Install dependencies:
```bash
npm install
```

2) Run the server:
```bash
npm run start
```

Open `http://localhost:3000` in your browser:
- Paste JSON into the textarea
- Upload a `.xlsx` template containing placeholders like `{{name}}`
- Click Export to download the merged file

## Library Usage

```js
import { mergeExcelTemplate } from './lib/index.js';
import fs from 'fs/promises';

const template = await fs.readFile('./template.xlsx');
const data = { name: 'Stanley', order: { total: 199.99 } };
const result = await mergeExcelTemplate(template, data);
await fs.writeFile('./output.xlsx', result);
```

### Placeholder Notes
- Supports dotted and bracket paths:
  - `{{user.name}}`
  - `{{items[0].price}}`
- Missing values default to empty string. You can change with `onMissing` option.
- Non-primitive values are stringified with `JSON.stringify` by default.

## Options

```js
await mergeExcelTemplate(template, data, {
  onMissing: 'empty', // 'empty' | 'keep' | (expr) => string
  valueToString: (v) => String(v) // custom stringify
});
```

## Template Tips
- Put placeholders directly in cells (any sheet)
- Example cells:
  - `A1: Hello {{name}}`
  - `B3: Order Total: {{order.total}}`

## Scripts
- `npm start` — start the Express server
- `npm run dev` — start with nodemon (auto-reload)

## License
MIT


