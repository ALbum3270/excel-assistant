import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");

test("Enter submits, but IME candidate confirmation does not", () => {
  const start = source.indexOf('$input.addEventListener("keydown"');
  const end = source.indexOf("// Theme —", start);
  assert.ok(start >= 0 && end > start, "taskpane keydown handler source anchors must exist");

  let handler;
  const submissions = [];
  const sandbox = {
    $input: {
      addEventListener(name, fn) {
        assert.equal(name, "keydown");
        handler = fn;
      },
    },
    $composer: { dispatchEvent: (event) => submissions.push(event.type) },
    Event: class Event {
      constructor(type) {
        this.type = type;
      }
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(start, end), sandbox, {
    filename: "taskpane-keydown-under-test.js",
  });

  const invoke = (overrides = {}) => {
    let prevented = false;
    handler({
      key: "Enter",
      shiftKey: false,
      isComposing: false,
      keyCode: 13,
      preventDefault: () => {
        prevented = true;
      },
      ...overrides,
    });
    return prevented;
  };

  assert.equal(invoke(), true);
  assert.deepEqual(submissions, ["submit"]);

  assert.equal(invoke({ isComposing: true, keyCode: 229 }), false);
  assert.equal(invoke({ isComposing: false, keyCode: 229 }), false);
  assert.equal(invoke({ shiftKey: true }), false);
  assert.equal(invoke({ key: "a", keyCode: 65 }), false);
  assert.deepEqual(submissions, ["submit"]);
});
