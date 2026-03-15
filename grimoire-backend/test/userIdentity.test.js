import assert from "node:assert/strict";
import test from "node:test";

import {
  extractAuthenticatedUser,
  requireAuthenticatedUser,
} from "../src/middleware/userIdentity.js";

function buildPrincipalHeader(payload) {
  return Buffer.from(JSON.stringify(payload), "utf8").toString("base64");
}

test("extractAuthenticatedUser resolves object id and preferred username from Easy Auth principal", () => {
  const request = {
    headers: new Headers({
      "x-ms-client-principal": buildPrincipalHeader({
        userDetails: "someone@example.com",
        claims: [
          {
            typ: "http://schemas.microsoft.com/identity/claims/objectidentifier",
            val: "ABCDEF12-3456-7890-abcd-ef1234567890",
          },
          {
            typ: "preferred_username",
            val: "someone@example.com",
          },
        ],
      }),
    }),
  };

  const user = extractAuthenticatedUser(request);

  assert.deepEqual(user, {
    objectId: "abcdef12-3456-7890-abcd-ef1234567890",
    upnOrEmail: "someone@example.com",
  });
});

test("extractAuthenticatedUser falls back to Easy Auth identity headers when claims are sparse", () => {
  const request = {
    headers: new Headers({
      "x-ms-client-principal": buildPrincipalHeader({ claims: [] }),
      "x-ms-client-principal-id": "12345678-1234-1234-1234-1234567890AB",
      "x-ms-client-principal-name": "fallback@example.com",
    }),
  };

  const user = extractAuthenticatedUser(request);

  assert.deepEqual(user, {
    objectId: "12345678-1234-1234-1234-1234567890ab",
    upnOrEmail: "fallback@example.com",
  });
});

test("requireAuthenticatedUser rejects requests without an authenticated principal", () => {
  const result = requireAuthenticatedUser({ headers: new Headers() }, {
    "Access-Control-Allow-Origin": "https://contoso.sharepoint.com",
  });

  assert.equal(result.authenticated, false);
  assert.equal(result.errorResponse?.status, 401);
  assert.equal(
    result.errorResponse?.jsonBody?.error,
    "Authenticated user identity required."
  );
});
