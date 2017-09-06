# Build an Add-in with Angular

### Step 1. Generate the Angular project by **Angular CLI**

If you never install [Angular CLI](https://github.com/angular/angular-cli) before, first install it globally.

```bash
npm install -g @angular/cli
```

Then generate your Angular app by

```bash
ng new my-addin
```

### Step 2. Generate the manifest file by **Office Toolbox**

If you never installed [Office Toolbox](https://needupdate) before, first install it globally.

```bash
npm install -g office-toolbox
```

If you installed it before, go to your app folder.

```bash
cd my-addin
```

Generate the manifest file following the steps below.

```bash
office-toolbox
```

![Generate](./img/office-toolbox-generate.png)

You should be able to see your manifest file with the name ends with **manifest.xml**.

### Step 3. Initialize

Open **src/index.html**, add

```html
<script src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.debug.js"></script>
```

before `</head>` tag.

Open **src/main.ts**, add `Office.initialize` out of `platformBrowserDynamic().bootstrapModule(AppModule);` like below:

```typescript
declare const Office: any;

Office.initialize = () => {
  platformBrowserDynamic().bootstrapModule(AppModule);
};
```

### Step 4. Add "Color Me"

Open **src/app/app.component.html**. Replace by

```html
<button (click)="onColorMe()">Color Me</button>
```

Open **src/app/app.component.ts**. Replace by

```typescript
import { Component } from '@angular/core';

declare const Excel: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  onColorMe() {
    Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }
}
```

### Step 5. Run the app

Run the dev server through the terminal.

```bash
npm start
```

or

```bash
ng serve
```

### Step 6. Side load the manifest file by **Office Toolbox**

To run the add-in, you need side-load the add-in in the Excel.

Run this in terminal and following the steps below.

```bash
office-toolbox
```

![Sideload](./img/office-toolbox-sideload.png)

Open Excel and click your add-in to load.

Congratulations you just finish your first Angular add-in for Excel!

