# xlsx-export

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENCE) [![GitHub Sponsors](https://img.shields.io/badge/Sponsor-GitHub%20Sponsors-blue?logo=github)](https://github.com/sponsors/your-name)

Angular 21 compatible package for generating Excel `.xlsx` files without external dependencies.

## Features

- Exports a single worksheet
- Builds an Open XML `.xlsx` archive using TypeScript and browser APIs
- Uses object keys as column headers and values as row content
- Avoids external libraries for portability

## Usage

```ts
import { ExcelExportService } from 'xls-export';

const data = [
  { Name: 'Ana', Age: 28, Active: true },
  { Name: 'Luis', Age: 32, Active: false },
];

const service = new ExcelExportService();
const blob = service.generate(data, { filename: 'clients.xlsx', sheetName: 'Sheet 1' });

// Download in the browser
service.save(data, { filename: 'clients.xlsx' });
```

## Angular Example

```ts
import { Component } from '@angular/core';
import { ExcelExportService } from 'xls-export';

@Component({
  selector: 'app-export-demo',
  template: `
    <button (click)="export()">Export Excel</button>
  `,
})
export class ExportDemoComponent {
  private readonly data = [
    { Name: 'Ana', Age: 28, Active: true },
    { Name: 'Luis', Age: 32, Active: false },
  ];

  export(): void {
    const service = new ExcelExportService();
    service.save(this.data, { filename: 'clients.xlsx', sheetName: 'Sheet 1' });
  }
}
```

## API

- `generate(rows, options)` → `Blob`
- `save(rows, options)` → downloads the spreadsheet as a file in the browser

## Support

If you find this project useful, consider supporting its development:

- GitHub Sponsors: https://github.com/sponsors/your-name
- Buy Me a Coffee: https://buymeacoffee.com/your-name
- PayPal: https://paypal.me/your-name

## Build

```bash
npm install
npm run build
```
