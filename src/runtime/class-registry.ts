import type { VbValue } from './values.ts';
import { VbEmpty } from './values.ts';

export type VbPropertyGetter = () => VbValue;
export type VbPropertySetter = (value: VbValue) => void;

export interface VbProperty {
  name: string;
  get?: VbPropertyGetter;
  let?: VbPropertySetter;
  set?: VbPropertySetter;
}

export interface VbMethodInfo {
  name: string;
  func: (...args: VbValue[]) => VbValue;
  isSub: boolean;
}

export class VbClass {
  constructor(
    public name: string,
    public properties: Map<string, VbProperty> = new Map(),
    public methods: Map<string, VbMethodInfo> = new Map(),
    public initializer?: (instance: VbObjectInstance) => void
  ) {}
}

export class VbObjectInstance {
  private properties: Map<string, VbValue> = new Map();
  private propertyAccessors: Map<string, VbProperty> = new Map();

  constructor(
    public classInfo: VbClass,
    public prototype: VbObjectInstance | null = null
  ) {
    if (classInfo.initializer) {
      classInfo.initializer(this);
    }
    classInfo.properties.forEach((prop, name) => {
      this.propertyAccessors.set(name.toLowerCase(), prop);
      if (!prop.get && !prop.let && !prop.set) {
        this.properties.set(name.toLowerCase(), VbEmpty);
      }
    });
  }

  getProperty(name: string): VbValue {
    const lowerName = name.toLowerCase();
    const accessor = this.propertyAccessors.get(lowerName);
    if (accessor?.get) {
      return accessor.get.call(this);
    }
    if (accessor && !accessor.get && !accessor.let && !accessor.set) {
      const value = this.properties.get(lowerName);
      if (value !== undefined) {
        return value;
      }
    }
    const value = this.properties.get(lowerName);
    if (value !== undefined) {
      return value;
    }
    if (this.prototype) {
      return this.prototype.getProperty(name);
    }
    return VbEmpty;
  }

  setProperty(name: string, value: VbValue, isSet: boolean = false): void {
    const lowerName = name.toLowerCase();
    const accessor = this.propertyAccessors.get(lowerName);
    if (accessor) {
      if (isSet && accessor.set) {
        accessor.set.call(this, value);
        return;
      }
      if (!isSet && accessor.let) {
        accessor.let.call(this, value);
        return;
      }
      if (!accessor.get && !accessor.let && !accessor.set) {
        this.properties.set(lowerName, value);
        return;
      }
    }
    this.properties.set(lowerName, value);
  }

  hasProperty(name: string): boolean {
    const lowerName = name.toLowerCase();
    const has = this.properties.has(lowerName) || 
           this.propertyAccessors.has(lowerName) ||
           (this.prototype?.hasProperty(name) ?? false);
    return has;
  }

  getMethod(name: string): VbMethodInfo | undefined {
    return this.classInfo.methods.get(name.toLowerCase());
  }

  hasMethod(name: string): boolean {
    return this.classInfo.methods.has(name.toLowerCase());
  }
}

export class VbClassRegistry {
  private classes: Map<string, VbClass> = new Map();

  register(cls: VbClass): void {
    this.classes.set(cls.name.toLowerCase(), cls);
  }

  registerClass(name: string, creator: () => VbValue): void {
    const cls = new VbClass(name);
    cls.initializer = (instance: VbObjectInstance) => {
      const value = creator();
      if (value.type === 'Object' && value.value) {
        const obj = value.value as Record<string, unknown>;
        instance.properties = new Map(Object.entries(obj).map(([k, v]) => [k.toLowerCase(), v as VbValue]));
      }
    };
    this.classes.set(name.toLowerCase(), cls);
  }

  get(name: string): VbClass | undefined {
    return this.classes.get(name.toLowerCase());
  }

  has(name: string): boolean {
    return this.classes.has(name.toLowerCase());
  }

  createInstance(name: string, _args: VbValue[]): VbObjectInstance {
    const cls = this.classes.get(name.toLowerCase());
    if (!cls) {
      throw new Error(`Undefined class: ${name}`);
    }

    const instance = new VbObjectInstance(cls);

    const initMethod = cls.methods.get('class_initialize');
    if (initMethod) {
      initMethod.func.call(instance);
    }

    const initProperty = cls.properties.get('class_initialize');
    if (initProperty && initProperty.get) {
      initProperty.get.call(instance);
    }

    return instance;
  }
}
