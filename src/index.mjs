import { DurableObject } from "cloudflare:workers";

const API_PORT = 8080;
const API_HEALTH_URL = "http://container/api/health";
const INACTIVITY_TIMEOUT_MS = 5 * 60 * 1000;

export class JjmApiContainer extends DurableObject {
  starting;

  async fetch(request) {
    this.starting ??= this.startAndWaitForPort().finally(() => {
      this.starting = undefined;
    });
    await this.starting;

    const url = new URL(request.url);
    url.protocol = "http:";
    url.host = "container";
    const forwarded = new Request(url, request);
    forwarded.headers.delete("host");
    return this.ctx.container.getTcpPort(API_PORT).fetch(forwarded);
  }

  async startAndWaitForPort() {
    const container = this.ctx.container;
    if (!container) throw new Error("Cloudflare no proporcionó el contenedor de JJM Clock.");

    if (!this.env.SUPABASE_URL || !this.env.SUPABASE_SECRET_KEY) {
      throw new Error("Configura SUPABASE_URL y el secreto SUPABASE_SECRET_KEY en el Worker.");
    }
    if (!this.env.FRONTEND_ORIGINS) {
      throw new Error("Configura FRONTEND_ORIGINS en el Worker.");
    }

    if (!container.running) {
      container.start({
        image: container.images.base,
        instance: "lite",
        enableInternet: true,
        env: {
          PORT: String(API_PORT),
          NODE_ENV: "production",
          SUPABASE_URL: this.env.SUPABASE_URL,
          SUPABASE_SECRET_KEY: this.env.SUPABASE_SECRET_KEY,
          FRONTEND_ORIGINS: this.env.FRONTEND_ORIGINS,
          COOKIE_SAME_SITE: this.env.COOKIE_SAME_SITE || "lax",
          VAPID_PUBLIC_KEY: this.env.VAPID_PUBLIC_KEY || "",
          VAPID_PRIVATE_KEY: this.env.VAPID_PRIVATE_KEY || "",
          VAPID_SUBJECT: this.env.VAPID_SUBJECT || "mailto:admin@gruponerba.com"
        }
      });
    }

    await container.setInactivityTimeout(INACTIVITY_TIMEOUT_MS);
    const port = container.getTcpPort(API_PORT);
    let lastError;
    for (let attempt = 0; attempt < 100; attempt++) {
      try {
        const response = await port.fetch(API_HEALTH_URL, {
          signal: AbortSignal.timeout(1000)
        });
        await response.body?.cancel();
        if (!response.ok) throw new Error("Health check respondió " + response.status);
        return;
      } catch (error) {
        lastError = error;
        await new Promise(resolve => setTimeout(resolve, 200));
      }
    }
    throw new Error("La API de JJM Clock no respondió en el puerto " + API_PORT, {
      cause: lastError
    });
  }
}

export default {
  fetch(request, env) {
    return env.API.getByName("jjmclock-api").fetch(request);
  }
};
