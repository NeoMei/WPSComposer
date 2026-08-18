(function () {
  "use strict";

  const ERROR_CODES = [
    "CONVERSION_COMMAND_FAILED",
    "GENERATION_COMMAND_FAILED",
    "OPERATION_PLAN_INVALID",
    "INTERACTIVE_INPUT_REQUIRED",
    "NO_VISIBLE_WORKSHEETS"
  ];

  let started = false;

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  async function readJson(response) {
    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }

  async function request(session, path, options) {
    const headers = Object.assign({}, options.headers || {}, {
      "Content-Type": "application/json"
    });
    if (session.token) {
      headers.Authorization = `Bearer ${session.token}`;
    }
    const response = await fetch(
      `${session.bridgeUrl}${path}`,
      Object.assign({}, options, {headers})
    );
    if (!response.ok && response.status !== 204) {
      throw new Error(`Bridge ${path} failed with HTTP ${response.status}`);
    }
    return {status: response.status, body: await readJson(response)};
  }

  async function sendResult(session, result) {
    await request(session, "/v1/result", {
      method: "POST",
      body: JSON.stringify(Object.assign({}, result, {
        component: session.component,
        clientId: session.clientId
      }))
    });
  }

  async function run() {
    const sessionResponse = await fetch("./session.json", {cache: "no-store"});
    if (!sessionResponse.ok) {
      throw new Error("session.json is unavailable");
    }
    const bootstrap = await sessionResponse.json();
    const capability = bootstrap.capability || "";
    if (!capability) {
      throw new Error("Bridge bootstrap capability is unavailable");
    }
    const claimed = await request(bootstrap, "/v1/session", {
      method: "POST",
      body: JSON.stringify({
        component: bootstrap.component,
        clientId: bootstrap.clientId,
        capability
      })
    });
    const token = claimed.body.token;
    const session = Object.assign({}, bootstrap, {token});
    await request(session, "/v1/register", {
      method: "POST",
      body: JSON.stringify({
        component: session.component,
        clientId: session.clientId
      })
    });
    if (bootstrap.activationFixture &&
        typeof window.WPSComposerProbe.closeActivationFixture === "function") {
      window.WPSComposerProbe.closeActivationFixture(
        bootstrap.activationFixture
      );
    }

    let failures = 0;
    while (true) {
      let next;
      try {
        next = await request(
          session,
          `/v1/next?component=${encodeURIComponent(session.component)}` +
            `&clientId=${encodeURIComponent(session.clientId)}`,
          {method: "GET"}
        );
        failures = 0;
      } catch (pollError) {
        // A bearer belongs to one bridge session. Never replay the consumed
        // bootstrap capability after a restart; fail closed and let the host
        // activate a newly registered profile.
        if (String(pollError && pollError.message).indexOf("HTTP 401") !== -1) {
          throw pollError;
        }
        failures += 1;
        await sleep(Math.min(500 * failures, 5000));
        continue;
      }
      if (next.status === 204) {
        continue;
      }
      const command = next.body;
      try {
        const value = await window.WPSComposerProbe.handleCommand(command);
        await sendResult(session, {
          id: command.id,
          ok: true,
          value,
          error: null
        });
      } catch (error) {
        try {
          await sendResult(session, {
            id: command.id,
            ok: false,
            value: {},
            error: {
              code: error && ERROR_CODES.indexOf(error.code) !== -1
                ? error.code
                : "CONVERSION_COMMAND_FAILED",
              message: String(error && error.message ? error.message : error),
              stack: String(error && error.stack ? error.stack : "")
            }
          });
        } catch (reportError) {
          console.error("Failed to report command result", reportError);
          await sleep(500);
        }
      }
    }
  }

  window.OnAddinLoad = function () {
    if (!started) {
      started = true;
      run().catch(function (error) {
        console.error("WPSComposer Phase 0 add-in failed", error);
      });
    }
    return true;
  };
}());
