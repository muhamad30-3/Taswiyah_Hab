import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const pnpm = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
const windows = process.platform === "win32";

function runPnpm(args, cwd) {
  return spawn(
    windows ? process.env.ComSpec ?? "cmd.exe" : pnpm,
    windows ? ["/d", "/s", "/c", `${pnpm} ${args.join(" ")}`] : args,
    { cwd, stdio: "inherit" },
  );
}

const services = [
  { name: "Web", child: runPnpm(["run", "dev:web"], root) },
  { name: "API", child: runPnpm(["--dir", "artifacts/api-server", "run", "dev"], root) },
];

let stopping = false;

function stopChildren(exitCode = 0) {
  if (stopping) return;
  stopping = true;
  for (const { child } of services) {
    if (child.exitCode === null && !child.killed) child.kill();
  }
  process.exitCode = exitCode;
}

for (const { name, child } of services) {
  child.on("error", error => {
    console.error(`Could not start the ${name} development service:`, error.message);
    stopChildren(1);
  });
  child.on("exit", (code, signal) => {
    if (stopping) return;
    const exitCode = code ?? (signal ? 1 : 0);
    console.error(`${name} development service exited (${signal ?? `code ${exitCode}`}); stopping the other service.`);
    stopChildren(exitCode);
  });
}

process.on("SIGINT", () => stopChildren(0));
process.on("SIGTERM", () => stopChildren(0));
