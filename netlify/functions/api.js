const serverless = require("serverless-http");

// server.js exports the Express app (and only listens when run directly).
const app = require("../../server");

const basePath = "/.netlify/functions/api";
const handler = serverless(app);

module.exports.handler = async (event, context) => {
  // Make routing predictable for Express.
  // With redirects, Netlify may invoke the function with a full path like:
  //   /.netlify/functions/api/api/summary
  // Express expects:
  //   /api/summary
  if (event && typeof event.path === "string" && event.path.startsWith(basePath)) {
    const next = event.path.slice(basePath.length) || "/";
    event.path = next.startsWith("/") ? next : `/${next}`;
  }
  return handler(event, context);
};
