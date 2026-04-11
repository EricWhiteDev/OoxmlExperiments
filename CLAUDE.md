# CLAUDE.md

## Code Style

- Always use curly braces for `if`, `else`, `for`, `while`, and other control flow statements. Never write single-line conditionals without braces.
- Always put the conditional body on its own line, not on the same line as the `if`.
- Use "one true brace style" (1tbs): opening brace on the same line as the statement, closing brace on its own line.
- These rules are enforced by ESLint (`curly: error`, `brace-style: ["error", "1tbs"]`) and should never be disabled.

Example:
```typescript
// Correct
if (!stylePart) {
  return null;
}

// Wrong
if (!stylePart) return null;
```

## Build

- `npm run build:dev` to build in development mode
- `npm run dev-server` to start the webpack dev server on https://localhost:3000
- `npm run start` to launch the add-in in Word
