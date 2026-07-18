(function () {
  "use strict";

  let started = false;

  async function readJson(response) {
    const text = await response.text();
    return text ? JSON.parse(text) : null;
  }

  async function request(session, path, options) {
    const headers = Object.assign({}, options.headers || {}, {
      Authorization: `Bearer ${session.token}`,
      "Content-Type": "application/json"
    });
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
      body: JSON.stringify(result)
    });
  }

  async function run() {
    const sessionResponse = await fetch("./session.json", {cache: "no-store"});
    if (!sessionResponse.ok) {
      throw new Error("session.json is unavailable");
    }
    const session = await sessionResponse.json();
    await request(session, "/v1/register", {
      method: "POST",
      body: JSON.stringify({component: session.component})
    });

    while (true) {
      const next = await request(
        session,
        `/v1/next?component=${encodeURIComponent(session.component)}`,
        {method: "GET"}
      );
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
        await sendResult(session, {
          id: command.id,
          ok: false,
          value: {},
          error: {
            code: error && typeof error.code === "string"
              ? error.code
              : "CONVERSION_COMMAND_FAILED",
            message: String(error && error.message ? error.message : error),
            stack: String(error && error.stack ? error.stack : "")
          }
        });
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
