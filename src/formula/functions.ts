import "./math";
import "./text";
import "./dates";
import "./lookup";
import "./finance";
import { functions } from "./registry";
import { signatures } from "./signatures";
export { functions };
export const listFunctions = () =>
  [...functions.values()]
    .map(({ evaluate, ...metadata }) => ({
      ...metadata,
      signature:
        metadata.signature ??
        signatures[metadata.name] ??
        (metadata.maxArgs === 0
          ? ""
          : metadata.maxArgs === 1
            ? "value"
            : "value1, ..."),
    }))
    .sort((a, b) => a.name.localeCompare(b.name));
