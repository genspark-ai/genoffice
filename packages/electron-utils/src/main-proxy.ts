import { EnvHttpProxyAgent, type Dispatcher } from 'undici'

/**
 * The NO_PROXY list the main-process dispatcher ships with: the caller's own
 * entries first, then loopback. undici matches entries exactly (or by `*.` or
 * leading-dot suffix), and the URL host of IPv6 loopback keeps its brackets.
 */
export function mergeNoProxy(userNoProxy: string | undefined): string {
  return [userNoProxy, 'localhost', '127.0.0.1', '[::1]'].filter((v) => !!v).join(',')
}

/**
 * The dispatcher behind the main process's global fetch when a proxy is in
 * play. A plain ProxyAgent sends loopback targets to the proxy too — as
 * absolute-form requests for http:// — so a Custom provider on 127.0.0.1
 * (Ollama, LM Studio, a local relay) gets routed to a remote node and comes
 * back with a gateway page (genoffice#2084). EnvHttpProxyAgent consults NO_PROXY per
 * request; the loopback defaults ride along even when the user set none.
 *
 * The resolved proxyUrl is passed for both schemes, matching what the
 * ProxyAgent it replaces did; the caller keeps the env/resolveProxy probing.
 */
export async function createMainProcessProxyDispatcher(proxyUrl: string): Promise<Dispatcher> {
  return new EnvHttpProxyAgent({
    httpProxy: proxyUrl,
    httpsProxy: proxyUrl,
    noProxy: mergeNoProxy(process.env.no_proxy ?? process.env.NO_PROXY),
  })
}
