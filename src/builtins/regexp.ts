import type { VbValue } from '../runtime/index.ts';

interface VbMatch {
  firstIndex: number;
  length: number;
  value: string;
  subMatches: string[];
}

interface VbRegExpState {
  pattern: string;
  global: boolean;
  ignoreCase: boolean;
  multiline: boolean;
}

function createMatchObject(match: VbMatch): VbValue {
  return {
    type: 'Object',
    value: {
      classInfo: { name: 'Match' },
      getProperty: (name: string): VbValue => {
        const lowerName = name.toLowerCase();
        if (lowerName === 'firstindex') {
          return { type: 'Long', value: match.firstIndex };
        }
        if (lowerName === 'length') {
          return { type: 'Long', value: match.length };
        }
        if (lowerName === 'value') {
          return { type: 'String', value: match.value };
        }
        return { type: 'Empty', value: undefined };
      },
      hasProperty: (name: string): boolean => {
        const lowerName = name.toLowerCase();
        return ['firstindex', 'length', 'value'].includes(lowerName);
      },
      getPropertyNames: (): string[] => ['FirstIndex', 'Length', 'Value'],
    },
  };
}

function createSubMatchesObject(subMatches: string[]): VbValue {
  return {
    type: 'Object',
    value: {
      classInfo: { name: 'SubMatches' },
      getProperty: (name: string): VbValue => {
        const index = parseInt(name, 10);
        if (!isNaN(index) && index >= 0 && index < subMatches.length) {
          return { type: 'String', value: subMatches[index] ?? '' };
        }
        return { type: 'Empty', value: undefined };
      },
      hasProperty: (name: string): boolean => {
        const index = parseInt(name, 10);
        return !isNaN(index) && index >= 0 && index < subMatches.length;
      },
      hasMethod: (name: string): boolean => {
        return name.toLowerCase() === 'count';
      },
      getMethod: (name: string): { func: (...args: VbValue[]) => VbValue } => {
        if (name.toLowerCase() === 'count') {
          return { func: () => ({ type: 'Long', value: subMatches.length }) };
        }
        throw new Error(`Unknown SubMatches method: ${name}`);
      },
    },
  };
}

function createMatchesCollection(matches: VbMatch[]): VbValue {
  const matchObjects = matches.map(createMatchObject);
  
  return {
    type: 'Object',
    value: {
      classInfo: { name: 'Matches' },
      getProperty: (name: string): VbValue => {
        const index = parseInt(name, 10);
        if (!isNaN(index) && index >= 0 && index < matchObjects.length) {
          return matchObjects[index];
        }
        return { type: 'Empty', value: undefined };
      },
      hasProperty: (name: string): boolean => {
        const index = parseInt(name, 10);
        return !isNaN(index) && index >= 0 && index < matchObjects.length;
      },
      hasMethod: (name: string): boolean => {
        return name.toLowerCase() === 'count' || name.toLowerCase() === 'item';
      },
      getMethod: (name: string): { func: (...args: VbValue[]) => VbValue } => {
        const lowerName = name.toLowerCase();
        if (lowerName === 'count') {
          return { func: () => ({ type: 'Long', value: matches.length }) };
        }
        if (lowerName === 'item') {
          return {
            func: (index: VbValue): VbValue => {
              const idx = Math.floor(Number(index.value ?? 0));
              if (idx >= 0 && idx < matchObjects.length) {
                return matchObjects[idx];
              }
              return { type: 'Empty', value: undefined };
            },
          };
        }
        throw new Error(`Unknown Matches method: ${name}`);
      },
    },
  };
}

function createRegExpObject(state: VbRegExpState): VbValue {
  return {
    type: 'Object',
    value: {
      classInfo: { name: 'RegExp' },
      regexpState: state,
      getProperty: (name: string): VbValue => {
        const lowerName = name.toLowerCase();
        if (lowerName === 'pattern') {
          return { type: 'String', value: state.pattern };
        }
        if (lowerName === 'global') {
          return { type: 'Boolean', value: state.global };
        }
        if (lowerName === 'ignorecase') {
          return { type: 'Boolean', value: state.ignoreCase };
        }
        if (lowerName === 'multiline') {
          return { type: 'Boolean', value: state.multiline };
        }
        return { type: 'Empty', value: undefined };
      },
      setProperty: (name: string, value: VbValue): void => {
        const lowerName = name.toLowerCase();
        if (lowerName === 'pattern') {
          state.pattern = String(value.value ?? '');
        } else if (lowerName === 'global') {
          state.global = Boolean(value.value);
        } else if (lowerName === 'ignorecase') {
          state.ignoreCase = Boolean(value.value);
        } else if (lowerName === 'multiline') {
          state.multiline = Boolean(value.value);
        }
      },
      hasProperty: (name: string): boolean => {
        const lowerName = name.toLowerCase();
        return ['pattern', 'global', 'ignorecase', 'multiline'].includes(lowerName);
      },
      hasMethod: (name: string): boolean => {
        return ['execute', 'test', 'replace'].includes(name.toLowerCase());
      },
      getMethod: (name: string): { func: (...args: VbValue[]) => VbValue } => {
        const lowerName = name.toLowerCase();
        
        if (lowerName === 'test') {
          return {
            func: (sourceString: VbValue): VbValue => {
              const source = String(sourceString.value ?? '');
              if (!state.pattern) {
                return { type: 'Boolean', value: false };
              }
              try {
                const flags = (state.ignoreCase ? 'i' : '') + (state.multiline ? 'm' : '');
                const regex = new RegExp(state.pattern, flags);
                const result = regex.test(source);
                return { type: 'Boolean', value: result };
              } catch {
                return { type: 'Boolean', value: false };
              }
            },
          };
        }
        
        if (lowerName === 'execute') {
          return {
            func: (sourceString: VbValue): VbValue => {
              const source = String(sourceString.value ?? '');
              if (!state.pattern) {
                return createMatchesCollection([]);
              }
              try {
                const flags = (state.ignoreCase ? 'i' : '') + (state.multiline ? 'm' : '') + (state.global ? 'g' : '');
                const regex = new RegExp(state.pattern, flags);
                const matches: VbMatch[] = [];
                let match;
                
                while ((match = regex.exec(source)) !== null) {
                  matches.push({
                    firstIndex: match.index,
                    length: match[0].length,
                    value: match[0],
                    subMatches: match.slice(1).map(m => m ?? ''),
                  });
                  
                  if (!state.global) {
                    break;
                  }
                }
                
                return createMatchesCollection(matches);
              } catch {
                return createMatchesCollection([]);
              }
            },
          };
        }
        
        if (lowerName === 'replace') {
          return {
            func: (sourceString: VbValue, replaceString: VbValue): VbValue => {
              const source = String(sourceString.value ?? '');
              const replace = String(replaceString.value ?? '');
              if (!state.pattern) {
                return { type: 'String', value: source };
              }
              try {
                const flags = (state.ignoreCase ? 'i' : '') + (state.multiline ? 'm' : '') + (state.global ? 'g' : '');
                const regex = new RegExp(state.pattern, flags);
                const result = source.replace(regex, replace);
                return { type: 'String', value: result };
              } catch {
                return { type: 'String', value: source };
              }
            },
          };
        }
        
        throw new Error(`Unknown RegExp method: ${name}`);
      },
    },
  };
}

export function registerRegExp(context: { 
  functionRegistry: { register: (name: string, func: (...args: VbValue[]) => VbValue) => void };
  classRegistry: { registerClass: (name: string, creator: () => VbValue) => void };
}): void {
  context.classRegistry.registerClass('RegExp', () => {
    return createRegExpObject({
      pattern: '',
      global: false,
      ignoreCase: false,
      multiline: false,
    });
  });
  
  context.functionRegistry.register('RegExp', (): VbValue => {
    return createRegExpObject({
      pattern: '',
      global: false,
      ignoreCase: false,
      multiline: false,
    });
  });
}
