# Quick Start

This guide will help you get started with `@devscholar/vbs-engine-js` in minutes. We'll create two example projects: one for browser usage and another for Node.js.

---

## Browser Example

### 1. Create a New Project

```bash
mkdir vbs-browser-demo
cd vbs-browser-demo
npm init -y
```

### 2. Install the Package

```bash
npm install @devscholar/vbs-engine-js
```

### 3. Create an HTML File

Create `index.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>VBS Engine Browser Demo</title>
</head>
<body>
  <h1>VBScript Engine Demo</h1>
  <div id="output"></div>

  <script type="module">
    import { VBSInterpreter } from '@devscholar/vbs-engine-js';

    const interpreter = new VBSInterpreter();
    
    const vbsCode = `
      Dim message
      message = "Hello from VBScript!"
      MsgBox message
    `;

    try {
      interpreter.execute(vbsCode);
      document.getElementById('output').textContent = 'VBScript executed successfully!';
    } catch (error) {
      document.getElementById('output').textContent = 'Error: ' + error.message;
    }
  </script>
</body>
</html>
```

### 4. Run the Example

You can use any static file server. For example:

```bash
# Using Python
python -m http.server 8080

# Or using Node.js
npx serve .
```

Then open `http://localhost:8080` in your browser.

---

## Node.js Example

### 1. Create a New Project

```bash
mkdir vbs-node-demo
cd vbs-node-demo
npm init -y
```

### 2. Install the Package

```bash
npm install @devscholar/vbs-engine-js
```

### 3. Create a JavaScript File

Create `demo.js`:

```javascript
import { VBSInterpreter } from '@devscholar/vbs-engine-js';

const interpreter = new VBSInterpreter();

const vbsCode = `
  Dim x, y, result
  x = 10
  y = 20
  result = x + y
  
  Function Add(a, b)
    Add = a + b
  End Function
  
  Dim sum
  sum = Add(5, 3)
  
  ' Output will be available through the interpreter
`;

try {
  const result = interpreter.execute(vbsCode);
  console.log('VBScript executed successfully!');
  console.log('Result:', result);
} catch (error) {
  console.error('Error:', error.message);
}
```

### 4. Update package.json

Add `"type": "module"` to your `package.json`:

```json
{
  "name": "vbs-node-demo",
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "start": "node demo.js"
  },
  "dependencies": {
    "@devscholar/vbs-engine-js": "^0.0.6"
  }
}
```

### 5. Run the Example

```bash
node demo.js
```

---

## Next Steps

- Check out the [API documentation](../README.md) for more details
- Explore [compatibility features](./compatibility.md) for IE migration
- See the `examples/` folder in the repository for more complex use cases
