// Logic.gs 를 Node 에서 불러오기 (Apps Script 와 같은 전역 스코프 방식)
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
export function loadLogic() {
  const ctx = {};
  vm.createContext(ctx);
  vm.runInContext(readFileSync(new URL('../src/Logic.gs', import.meta.url), 'utf8'), ctx);
  return ctx;
}
