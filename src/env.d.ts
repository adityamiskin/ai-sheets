/// <reference types="vite/client" />

// Declare `process` for type-checking of convex files that use `process.env`
// when resolved through import chains from frontend code.
declare const process: {
  env: Record<string, string | undefined>;
};
