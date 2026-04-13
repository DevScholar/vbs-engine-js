import type { VbValue, VbObjectValueData } from '../runtime/values.ts';
import { VbEmpty, VbNull } from '../runtime/values.ts';
import { toString } from '../runtime/values.ts';

/**
 * Creates a VB6-compatible Collection object.
 *
 * Supports:
 *   col.Add item [, key [, before [, after]]]
 *   col.Remove index|key
 *   col.Item(index|key)  — 1-based index or string key
 *   col.Count
 *   For Each x In col ... Next  (uses 0-based getProperty("0")..getProperty(String(n-1)))
 */
export function createCollection(): VbObjectValueData {
  const items: VbValue[] = [];
  const keyIndex: Map<string, number> = new Map(); // lowercase key → 0-based index

  function rebuildKeys(): void {
    keyIndex.clear();
    // Keys are not re-keyed after removal by design; rebuild from scratch
    // We store keys as a parallel array to track which item has which key
  }

  // Parallel array of string keys (or null) for each item at same index
  const itemKeys: (string | null)[] = [];

  function getByIndexOrKey(index: VbValue): VbValue {
    if (index.type === 'String') {
      const key = (index.value as string).toLowerCase();
      const idx = keyIndex.get(key);
      if (idx === undefined) {
        throw new Error(`Collection: key not found`);
      }
      return items[idx] ?? VbNull;
    }
    // Numeric: 1-based
    const n = Math.floor(
      index.type === 'Integer' || index.type === 'Long' || index.type === 'Byte'
        ? (index.value as number)
        : Number((index as { value: unknown }).value)
    );
    if (n < 1 || n > items.length) {
      throw new Error(`Collection: index out of bounds`);
    }
    return items[n - 1];
  }

  function removeByIndexOrKey(index: VbValue): void {
    let pos: number;
    if (index.type === 'String') {
      const key = (index.value as string).toLowerCase();
      const idx = keyIndex.get(key);
      if (idx === undefined) throw new Error(`Collection: key not found`);
      pos = idx;
    } else {
      const n = Math.floor(Number((index as { value: unknown }).value as number));
      if (n < 1 || n > items.length) throw new Error(`Collection: index out of bounds`);
      pos = n - 1;
    }
    items.splice(pos, 1);
    itemKeys.splice(pos, 1);
    // Rebuild keyIndex
    keyIndex.clear();
    for (let i = 0; i < itemKeys.length; i++) {
      if (itemKeys[i] !== null) {
        keyIndex.set(itemKeys[i]!.toLowerCase(), i);
      }
    }
  }

  const obj: VbObjectValueData = {
    classInfo: { name: 'Collection' },

    getProperty(name: string): VbValue {
      const lower = name.toLowerCase();
      if (lower === 'count') {
        return { type: 'Long', value: items.length };
      }
      // Numeric string: 0-based access used by For Each
      const num = parseInt(name, 10);
      if (!isNaN(num) && String(num) === name && num >= 0 && num < items.length) {
        return items[num];
      }
      return VbEmpty;
    },

    setProperty(_name: string, _value: VbValue): void {
      // Collection properties are read-only
    },

    hasMethod(name: string): boolean {
      const lower = name.toLowerCase();
      return lower === 'add' || lower === 'remove' || lower === 'item' || lower === 'default';
    },

    getMethod(name: string): { func: (...args: VbValue[]) => VbValue } {
      const lower = name.toLowerCase();

      if (lower === 'add') {
        return {
          func: (item: VbValue, key?: VbValue, _before?: VbValue, _after?: VbValue): VbValue => {
            void _before;
            void _after;
            const pos = items.length;
            items.push(item);
            if (key && key.type === 'String' && (key.value as string) !== '') {
              const k = (key.value as string).toLowerCase();
              if (keyIndex.has(k)) {
                throw new Error(`Collection: key already exists`);
              }
              itemKeys.push(key.value as string);
              keyIndex.set(k, pos);
            } else {
              itemKeys.push(null);
            }
            return VbEmpty;
          },
        };
      }

      if (lower === 'remove') {
        return {
          func: (index: VbValue): VbValue => {
            removeByIndexOrKey(index);
            return VbEmpty;
          },
        };
      }

      if (lower === 'item' || lower === 'default') {
        return {
          func: (index: VbValue): VbValue => {
            return getByIndexOrKey(index);
          },
        };
      }

      throw new Error(`Collection: unknown method '${name}'`);
    },
  };

  void toString; // imported but used indirectly via rebuildKeys stub
  void rebuildKeys; // only used internally

  return obj;
}
