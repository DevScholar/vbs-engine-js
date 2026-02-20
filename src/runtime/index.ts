export * from './values.ts';
export * from './scope.ts';
export * from './function-registry.ts';
export * from './class-registry.ts';
export * from './vb-array.ts';
export * from './errors.ts';
export * from './context.ts';

// Performance optimization modules - these provide optimized alternatives
// Use them directly by importing from their specific modules
export * from './string-interner.ts';
export * from './value-pool.ts';
// Note: scope-optimized.ts, function-registry-optimized.ts, and context-optimized.ts
// provide optimized versions with the same class names. Import directly from those
// modules if you want to use the optimized versions.
