# Quick Start Guide

This guide will walk you through setting up and running vbs-engine-js examples from scratch.

## Prerequisites

- Node.js (v18 or higher recommended)
- npm

## Step 1: Navigate to vbs-engine-js

```bash
cd vbs-engine-js
```

## Step 2: Install Dependencies

```bash
npm install
```

## Step 3: Run Examples

### Browser Example

```bash
npm run dev
```

Then open http://localhost:5173/ in your browser.

### Node.js Example

```bash
node start.js examples/node-demo.ts
```

This will automatically build the project and run the example.

## Using start.js

The `start.js` script can run any TypeScript file in the project:

```bash
# Run a specific example
node start.js examples/node-demo.ts

# With custom runtime (node, bun, or deno)
node start.js examples/node-demo.ts --runtime=deno
```

## Troubleshooting

### Error: Cannot resolve dependency

Make sure dependencies are installed:
```bash
npm install
```

### Error: Module not found

The `start.js` script handles building automatically. If you encounter issues, try:
```bash
npm run build
```

### Error: Cannot find package 'tsx'

Install tsx as a dev dependency:
```bash
npm install --save-dev tsx
```

## Next Steps

- Read the [API documentation](../vbs-engine-js/README.md) for more features
- Explore the [examples](../vbs-engine-js/examples/) in the main repository
- Try exposing JavaScript objects to VBScript using `engine.addObject()`
