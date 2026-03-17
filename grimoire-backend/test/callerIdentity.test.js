import assert from "node:assert/strict";
import test from "node:test";

import { resolveCallerId } from "../src/middleware/callerIdentity.js";

function buildPrincipalHeader(payload) {
  return Buffer.from(JSON.stringify(payload), "utf8").toString("base64");
}

test("resolveCallerId returns easyauth:<objectId> when Easy Auth principal is present", () => {
  const request = {
    headers: new Headers({
      "x-ms-client-principal": buildPrincipalHeader({
        claims: [
          {
            typ: "http://schemas.microsoft.com/identity/claims/objectidentifier",
            val: "ABCDEF12-3456-7890-abcd-ef1234567890",
          },
        ],
      }),
    }),
  };

  const callerId = resolveCallerId(request);
  assert.equal(callerId, "easyauth:abcdef12-3456-7890-abcd-ef1234567890");
});

test("resolveCallerId returns anonymous when no Easy Auth identity is present", () => {
  const request = {
    headers: new Headers({}),
  };

  const callerId = resolveCallerId(request);
  assert.equal(callerId, "anonymous");
});

test("resolveCallerId returns anonymous when Easy Auth header is empty", () => {
  const request = {
    headers: new Headers({
      "x-ms-client-principal": "",
    }),
  };

  const callerId = resolveCallerId(request);
  assert.equal(callerId, "anonymous");
});

test("resolveCallerId returns anonymous when Easy Auth principal has no object ID", () => {
  const request = {
    headers: new Headers({
      "x-ms-client-principal": buildPrincipalHeader({
        claims: [
          { typ: "preferred_username", val: "someone@example.com" },
        ],
      }),
    }),
  };

  const callerId = resolveCallerId(request);
  assert.equal(callerId, "anonymous");
});
