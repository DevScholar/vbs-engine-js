import type { VbValue } from '../runtime/index.ts';

export function createObject(cls: VbValue, _servername?: VbValue): VbValue {
  const className = String(cls.value ?? cls);

  if (typeof window.ActiveXObject !== 'undefined') {
    try {
      const ax = new window.ActiveXObject(className);
      return { type: 'Object', value: {
        type: 'activex',
        object: ax,
        className
      }};
    } catch (e) {
      throw new Error(`ActiveX component can't create object: '${className}'`);
    }
  }

  throw new Error(`ActiveXObject is not supported in this browser environment. Cannot create: '${className}'`);
}

export function getObject(pathname?: VbValue, cls?: VbValue): VbValue {
  if (typeof window.ActiveXObject !== 'undefined') {
    const path = pathname ? String(pathname.value ?? pathname) : '';
    const className = cls ? String(cls.value ?? cls) : '';

    try {
      if (path) {
        const ax = new window.ActiveXObject(className || 'Scripting.FileSystemObject');
        return { type: 'Object', value: {
          type: 'activex',
          object: ax,
          className
        }};
      }
    } catch (e) {
      throw new Error(`ActiveX component can't create object: '${className || path}'`);
    }
  }

  const className = cls ? String(cls.value ?? cls) : '';
  throw new Error(`ActiveXObject is not supported in this browser environment. Cannot get: '${className}'`);
}
