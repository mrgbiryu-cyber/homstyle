import https from "node:https";

const productId = process.argv[2] || "G25110023371";
const origin = "https://homestyle.lge.co.kr";
const pageUrl = `${origin}/item?productId=${encodeURIComponent(productId)}`;

function fetchText(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { rejectUnauthorized: false }, (response) => {
        if (
          response.statusCode >= 300 &&
          response.statusCode < 400 &&
          response.headers.location
        ) {
          response.resume();
          resolve(fetchText(new URL(response.headers.location, url)));
          return;
        }
        response.setEncoding("utf8");
        let body = "";
        response.on("data", (chunk) => {
          body += chunk;
        });
        response.on("end", () => resolve(body));
      })
      .on("error", reject);
  });
}

const pageText = await fetchText(pageUrl);
const scriptPaths = [
  ...new Set(
    [...pageText.matchAll(/<script[^>]+src=["']([^"']+)["']/g)].map(
      (match) => match[1],
    ),
  ),
];

const terms =
  /setId|setID|set_id|set-id|setProductId|bundleId|packageId|combinationId/gi;
const results = (
  await Promise.all(
    scriptPaths.map(async (path) => {
      try {
        const text = await fetchText(new URL(path, origin));
        const matches = [...text.matchAll(terms)].slice(0, 20).map((match) => ({
          term: match[0],
          context: text
            .slice(Math.max(0, match.index - 180), match.index + 320)
            .replace(/\s+/g, " "),
        }));
        return matches.length ? { path, matches } : null;
      } catch (error) {
        return { path, error: String(error) };
      }
    }),
  )
).filter(Boolean);

console.log(
  JSON.stringify(
    {
      productId,
      pageUrl,
      scriptCount: scriptPaths.length,
      matchedScripts: results,
    },
    null,
    2,
  ),
);
