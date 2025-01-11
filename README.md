# Office-Addin-TaskPane-Svelte

This is Svelte template for Office Add-in project that can be managed with Visual Studio Code or any other editor. You can use it to create Office Add-ins for:

- Excel
- OneNote
- PowerPoint
- Project
- Word

## Usage

1. Fork this repository and open in any code editor to get started.

```
git clone https://github.com/krmanik/Office-Addin-TaskPane-Svelte.git
cd Office-Addin-TaskPane-Svelte

git checkout svelte-5
```

2. Install dependencies

```
npm install
```

3. Make changes to the project

- [src/](https://github.com/krmanik/Office-Addin-TaskPane-Svelte/tree/main/src)
- [manifest.xml](https://github.com/krmanik/Office-Addin-TaskPane-Svelte/tree/main/manifests)

4. Build the project

```
npm run build
```

5. Run the project and select office apps

```
npm run start
```

## Running the Generated Site

Launch the local HTTPS site on https://localhost:3000 by simply typing the following command in your console:

```
npm run dev
```

In another terminal run following commands to use Addin in Office apps.

```
npm start
```

### Excel Addin

```
npm run start:excel
```

### OneNote Addin

```
npm run start:onenote
```

### PowerPoint Addin

```
npm run start:powerpoint
```

### Project Addin

```
npm run start:project
```

### Word Addin

```
npm run start:word
```

> **Note:** Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run npm start, accept the prompt to install the certificate that the Yeoman generator provides.

Next, sideload the add-in in an Office application. See [Sideload an Office Add-in for testing](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).

## TypeScript

This template is written using [TypeScript](http://www.typescriptlang.org/).

## Debugging

This template supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Additional resources

* [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
* More Office Add-ins samples at [OfficeDev on Github](https://github.com/officedev)

* [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office)

* [Fluent-ui web components](https://learn.microsoft.com/en-us/fluent-ui/web-components/)
