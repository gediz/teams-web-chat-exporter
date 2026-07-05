// Registers a minimal ESM resolve hook so `node --test` can run tests
// that import project .ts modules using TypeScript-style extensionless
// relative imports (Node's type stripping requires explicit extensions).
// On a failed resolve of a relative extensionless specifier, retry with
// `.ts` appended. Used only by the pnpm test/check scripts.
import { register } from 'node:module';

register(new URL('data:text/javascript,' + encodeURIComponent(`
export async function resolve(specifier, context, nextResolve) {
  try {
    return await nextResolve(specifier, context);
  } catch (err) {
    if ((specifier.startsWith('./') || specifier.startsWith('../')) && !/\\.[a-z]+$/i.test(specifier)) {
      return nextResolve(specifier + '.ts', context);
    }
    throw err;
  }
}
`)));
