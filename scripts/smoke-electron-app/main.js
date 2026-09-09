const keepAlive = setInterval(() => {}, 1000);

process.once("exit", () => {
  clearInterval(keepAlive);
});

require("../smoke-conversions");
