"use strict";

const fetchImpl = globalThis.fetch || require("node-fetch");

const PAGE_SIZE = 1000;
const WRITE_BATCH_SIZE = 200;

class SupabaseStore {
  constructor({ url, key, collections }) {
    if (!url || !key) {
      throw new Error("Configura SUPABASE_URL y SUPABASE_SECRET_KEY en el backend.");
    }
    if (key.startsWith("sb_publishable_")) {
      throw new Error("El backend necesita una clave secret de Supabase, no una publishable.");
    }
    if (!key.startsWith("sb_secret_")) {
      try {
        const payload = JSON.parse(Buffer.from(key.split(".")[1], "base64url").toString("utf8"));
        if (payload.role !== "service_role") throw new Error("role");
      } catch {
        throw new Error("Usa una clave sb_secret_ o la clave antigua service_role de Supabase.");
      }
    }
    const parsedUrl = new URL(url);
    if (!["https:", "http:"].includes(parsedUrl.protocol)) {
      throw new Error("SUPABASE_URL debe comenzar con https:// (o http:// en desarrollo local).");
    }
    this.baseUrl = parsedUrl.toString().replace(/\/+$/, "");
    this.key = key;
    this.collections = [...collections];
    this.allowedCollections = new Set(collections);
  }

  async request(method, query = "", body = undefined, prefer = undefined) {
    const url = this.baseUrl + "/rest/v1/jjm_records" + query;
    const headers = {
      apikey: this.key,
      Accept: "application/json"
    };
    // Opaque sb_secret keys belong only in apikey; old service_role JWTs also
    // need Authorization for PostgREST to select the service_role role.
    if (!this.key.startsWith("sb_secret_")) {
      headers.Authorization = "Bearer " + this.key;
    }
    if (body !== undefined) headers["Content-Type"] = "application/json";
    if (prefer) headers.Prefer = prefer;

    const response = await fetchImpl(url, {
      method,
      headers,
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    const responseText = await response.text();
    if (!response.ok) {
      throw new Error(
        "Supabase respondió " + response.status + " al " + method +
        " jjm_records: " + responseText.slice(0, 500)
      );
    }
    if (!responseText) return null;
    try {
      return JSON.parse(responseText);
    } catch {
      throw new Error("Supabase devolvió una respuesta que no es JSON válido.");
    }
  }

  async load() {
    const db = Object.fromEntries(this.collections.map(name => [name, []]));
    const nextId = {};
    let offset = 0;

    while (true) {
      const query = new URLSearchParams({
        select: "collection,id,record",
        order: "collection.asc,id.asc",
        limit: String(PAGE_SIZE),
        offset: String(offset)
      });
      const rows = await this.request("GET", "?" + query.toString());
      if (!Array.isArray(rows)) {
        throw new Error("Supabase devolvió un formato inesperado al cargar jjm_records.");
      }

      for (const row of rows) {
        if (!this.allowedCollections.has(row.collection)) {
          throw new Error("Supabase contiene una colección desconocida: " + row.collection);
        }
        const id = Number(row.id);
        if (!Number.isSafeInteger(id) || id < 1 || !row.record || typeof row.record !== "object") {
          throw new Error("Supabase contiene un registro con formato inválido en " + row.collection + ".");
        }
        const record = { ...row.record, id };
        db[row.collection].push(record);
        nextId[row.collection] = Math.max(nextId[row.collection] || 1, id + 1);
      }

      if (rows.length < PAGE_SIZE) break;
      offset += rows.length;
    }
    return { db, nextId };
  }

  async persistChanges(currentDb, previousDb) {
    const upserts = [];
    const deletes = [];

    for (const collection of this.collections) {
      const currentRows = currentDb[collection] || [];
      const previousRows = previousDb[collection] || [];
      const previousById = new Map(previousRows.map(row => [String(row.id), row]));
      const currentIds = new Set();

      for (const row of currentRows) {
        const id = Number(row.id);
        if (!Number.isSafeInteger(id) || id < 1) {
          throw new Error("El registro de " + collection + " tiene un id no válido.");
        }
        const key = String(id);
        currentIds.add(key);
        const oldRow = previousById.get(key);
        if (!oldRow || JSON.stringify(oldRow) !== JSON.stringify(row)) {
          upserts.push({
            collection,
            id,
            record: row,
            updated_at: new Date().toISOString()
          });
        }
      }

      for (const oldRow of previousRows) {
        if (!currentIds.has(String(oldRow.id))) {
          deletes.push({ collection, id: Number(oldRow.id) });
        }
      }
    }

    for (let i = 0; i < upserts.length; i += WRITE_BATCH_SIZE) {
      await this.request(
        "POST",
        "?on_conflict=collection%2Cid",
        upserts.slice(i, i + WRITE_BATCH_SIZE),
        "resolution=merge-duplicates,return=minimal"
      );
    }

    for (let i = 0; i < deletes.length; i += WRITE_BATCH_SIZE) {
      const batch = deletes.slice(i, i + WRITE_BATCH_SIZE);
      const byCollection = new Map();
      for (const row of batch) {
        if (!byCollection.has(row.collection)) byCollection.set(row.collection, []);
        byCollection.get(row.collection).push(row.id);
      }
      for (const [collection, ids] of byCollection) {
        const query = new URLSearchParams({
          collection: "eq." + collection,
          id: "in.(" + ids.join(",") + ")"
        });
        await this.request("DELETE", "?" + query.toString(), undefined, "return=minimal");
      }
    }

    return { upserted: upserts.length, deleted: deletes.length };
  }
}

module.exports = { SupabaseStore };
