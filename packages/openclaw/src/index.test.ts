/**
 * Plugin-entry wiring: the Authorization Code flow needs a public HTTP route
 * that Microsoft redirects the browser to. Assert the entry registers it with
 * the right path, plugin-scoped auth, and exact matching (so only this path is
 * ever publicly reachable via Tailscale Funnel).
 */
import { describe, expect, it } from 'vitest';
import type {
  OpenClawPluginApi,
  OpenClawPluginHttpRouteParams,
} from 'openclaw/plugin-sdk/plugin-entry';

import outlookPlugin from './index.js';
import { AUTH_CALLBACK_PATH } from './auth-callback.js';

function captureRoutes(): {
  api: OpenClawPluginApi;
  routes: OpenClawPluginHttpRouteParams[];
} {
  const routes: OpenClawPluginHttpRouteParams[] = [];
  const api = {
    pluginConfig: {},
    registerTool() {},
    registerHttpRoute(params: OpenClawPluginHttpRouteParams) {
      routes.push(params);
    },
  } as unknown as OpenClawPluginApi;
  return { api, routes };
}

describe('outlook plugin entry — auth-callback route', () => {
  it('registers exactly one /outlook/auth-callback route', () => {
    const { api, routes } = captureRoutes();
    outlookPlugin.register(api);
    const matching = routes.filter((r) => r.path === AUTH_CALLBACK_PATH);
    expect(matching).toHaveLength(1);
  });

  it('mounts it with plugin auth, exact match, and a handler', () => {
    const { api, routes } = captureRoutes();
    outlookPlugin.register(api);
    const route = routes.find((r) => r.path === AUTH_CALLBACK_PATH);
    expect(route).toBeDefined();
    expect(route?.auth).toBe('plugin');
    expect(route?.match).toBe('exact');
    expect(typeof route?.handler).toBe('function');
  });
});
