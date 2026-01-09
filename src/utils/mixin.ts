type Constructor = Function & { prototype: any };

export default function mixin(
  targetClass: Constructor,
  mixinClasses: Array<{ ctor: Constructor; private?: boolean }>
) {
  mixinClasses.forEach((mixin) => {
    Object.getOwnPropertyNames(mixin.ctor.prototype).forEach((methodName) => {
      if (methodName !== "constructor") {
        if (mixin.private) {
          Object.defineProperty(targetClass.prototype, methodName, {
            value: mixin.ctor.prototype[methodName],
            writable: false,
            enumerable: false,
            configurable: false,
          });
        } else {
          targetClass.prototype[methodName] = mixin.ctor.prototype[methodName];
        }
      }
    });
  });
}
