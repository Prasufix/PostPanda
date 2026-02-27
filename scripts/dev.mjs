import { spawn } from "node:child_process";
import process from "node:process";

const npmCmd = process.platform === "win32" ? "npm.cmd" : "npm";

const children = [];
let shuttingDown = false;

function run(name, command, args) {
  const child = spawn(command, args, {
    cwd: process.cwd(),
    env: process.env,
    stdio: ["ignore", "pipe", "pipe"],
  });

  child.stdout.on("data", (chunk) => {
    process.stdout.write(prefixOutput(name, chunk.toString()));
  });

  child.stderr.on("data", (chunk) => {
    process.stderr.write(prefixOutput(name, chunk.toString()));
  });

  child.on("exit", (code) => {
    if (shuttingDown) return;
    shuttingDown = true;
    if (code && code !== 0) {
      console.error(`[${name}] beendet mit Exit-Code ${code}`);
    }
    shutdown(code ?? 0);
  });

  children.push(child);
}

function prefixOutput(name, text) {
  return text
    .split(/\r?\n/)
    .filter((line, index, list) => !(line.length === 0 && index === list.length - 1))
    .map((line) => `[${name}] ${line}`)
    .join("\n") + "\n";
}

function shutdown(code = 0) {
  for (const child of children) {
    if (!child.killed) {
      child.kill("SIGTERM");
    }
  }

  setTimeout(() => {
    for (const child of children) {
      if (!child.killed) {
        child.kill("SIGKILL");
      }
    }
    process.exit(code);
  }, 300);
}

process.on("SIGINT", () => {
  shuttingDown = true;
  shutdown(0);
});

process.on("SIGTERM", () => {
  shuttingDown = true;
  shutdown(0);
});

run("backend", "python3", ["main.py"]);
run("frontend", npmCmd, ["--prefix", "frontend", "run", "dev"]);
