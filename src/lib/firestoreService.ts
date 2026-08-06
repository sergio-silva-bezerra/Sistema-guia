import { getSupabaseClient } from './supabase';

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

// Local storage key helper
function getLocalKey(path: string, id?: string) {
  return id ? `nexus_sb_${path}_${id}` : `nexus_sb_${path}_list`;
}

// Local subscriber registry
interface Subscriber {
  path: string;
  coordinatorId?: string;
  callback: (data: any[]) => void;
}

const activeSubscribers = new Set<Subscriber>();

function notifySubscribers(path: string) {
  activeSubscribers.forEach(sub => {
    if (sub.path === path) {
      if (sub.coordinatorId) {
        firestoreService.getCollectionFiltered(path, sub.coordinatorId).then(sub.callback).catch(() => {});
      } else {
        firestoreService.getCollection(path).then(sub.callback).catch(() => {});
      }
    }
  });
}

// Get cached local list
function getLocalList<T>(path: string): T[] {
  try {
    const raw = localStorage.getItem(getLocalKey(path));
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return [];
}

// Save cached local list
function setLocalList(path: string, list: any[]) {
  try {
    localStorage.setItem(getLocalKey(path), JSON.stringify(list));
  } catch (e) {}
}

export const firestoreService = {
  async getCollection<T>(path: string): Promise<T[]> {
    const client = getSupabaseClient();
    const localItems = getLocalList<any>(path);

    if (client) {
      try {
        const { data, error } = await client
          .from('campaign_records')
          .select('record_id, payload')
          .eq('record_type', path);

        if (!error && data) {
          const supabaseItems = data.map(row => ({
            id: row.record_id,
            ...(row.payload || {})
          })) as any[];

          setLocalList(path, supabaseItems);
          if ((path === 'voters' || path === 'teams') && supabaseItems.length === 0) {
            import('./campaignSeedService').then(m => m.ensureSeedCampaignData()).catch(() => {});
          }
          return supabaseItems as T[];
        }
      } catch (e) {
        console.warn(`Supabase getCollection error for ${path}:`, e);
      }
    }
    if ((path === 'voters' || path === 'teams') && localItems.length === 0) {
      import('./campaignSeedService').then(m => m.ensureSeedCampaignData()).catch(() => {});
    }
    return localItems as T[];
  },

  async getDocument<T>(path: string, id: string): Promise<T | null> {
    const client = getSupabaseClient();
    if (client) {
      try {
        const { data, error } = await client
          .from('campaign_records')
          .select('record_id, payload')
          .eq('record_type', path)
          .eq('record_id', id)
          .maybeSingle();

        if (!error && data && data.payload) {
          return { id: data.record_id, ...data.payload } as T;
        }
      } catch (e) {
        console.warn(`Supabase getDocument error for ${path}/${id}:`, e);
      }
    }

    const localItems = getLocalList<any>(path);
    const found = localItems.find(item => item.id === id);
    if (found) return found as T;

    try {
      const raw = localStorage.getItem(getLocalKey(path, id));
      if (raw) return JSON.parse(raw);
    } catch (e) {}

    return null;
  },

  async setDocument(path: string, id: string, data: any, merge: boolean = true) {
    const client = getSupabaseClient();
    const payload = { id, ...data };
    const coordinatorId = data.coordinatorId || data.coordinator_id || data.userId || 'geral';

    // Local storage sync
    const items = getLocalList<any>(path);
    const idx = items.findIndex(item => item.id === id);
    if (idx >= 0) {
      items[idx] = merge ? { ...items[idx], ...payload } : payload;
    } else {
      items.push(payload);
    }
    setLocalList(path, items);
    try {
      localStorage.setItem(getLocalKey(path, id), JSON.stringify(payload));
    } catch (e) {}

    notifySubscribers(path);

    if (client) {
      try {
        const { data: existingRow } = await client
          .from('campaign_records')
          .select('id, payload')
          .eq('record_type', path)
          .eq('record_id', id)
          .maybeSingle();

        if (existingRow && existingRow.id) {
          const finalPayload = merge ? { ...(existingRow.payload || {}), ...payload } : payload;
          await client.from('campaign_records').update({
            coordinator_id: coordinatorId,
            payload: finalPayload
          }).eq('id', existingRow.id);
        } else {
          await client.from('campaign_records').insert({
            coordinator_id: coordinatorId,
            record_type: path,
            record_id: id,
            payload: payload
          });
        }
        notifySubscribers(path);
      } catch (e) {
        console.warn(`Supabase setDocument error for ${path}/${id}:`, e);
      }
    }
  },

  async updateDocument(path: string, id: string, data: any) {
    const existing = (await firestoreService.getDocument<any>(path, id)) || {};
    const updated = { ...existing, ...data };
    await firestoreService.setDocument(path, id, updated, true);
  },

  async deleteDocument(path: string, id: string) {
    const items = getLocalList<any>(path).filter(item => item.id !== id);
    setLocalList(path, items);
    try {
      localStorage.removeItem(getLocalKey(path, id));
    } catch (e) {}

    notifySubscribers(path);

    const client = getSupabaseClient();
    if (client) {
      try {
        await client
          .from('campaign_records')
          .delete()
          .eq('record_type', path)
          .eq('record_id', id);
        notifySubscribers(path);
      } catch (e) {
        console.warn(`Supabase deleteDocument error for ${path}/${id}:`, e);
      }
    }
  },

  async addDocument(path: string, data: any): Promise<string> {
    const id = data.id || `rec_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    await this.setDocument(path, id, data, true);
    return id;
  },

  subscribeToCollection<T>(path: string, callback: (data: T[]) => void) {
    const subObj = { path, callback: callback as any };
    activeSubscribers.add(subObj);

    // Initial emission
    firestoreService.getCollection<T>(path).then(callback);

    const client = getSupabaseClient();
    let channel: any = null;
    if (client) {
      channel = client
        .channel(`public:campaign_records:${path}_${Math.random().toString(36).substring(2, 7)}`)
        .on('postgres_changes', {
          event: '*',
          schema: 'public',
          table: 'campaign_records',
          filter: `record_type=eq.${path}`
        }, () => {
          firestoreService.getCollection<T>(path).then(callback);
        })
        .subscribe();
    }

    return () => {
      activeSubscribers.delete(subObj);
      if (client && channel) {
        client.removeChannel(channel);
      }
    };
  },

  subscribeToCollectionFiltered<T>(path: string, coordinatorId: string, callback: (data: T[]) => void) {
    const subObj = { path, coordinatorId, callback: callback as any };
    activeSubscribers.add(subObj);

    firestoreService.getCollectionFiltered<T>(path, coordinatorId).then(callback);

    const client = getSupabaseClient();
    let channel: any = null;
    if (client) {
      channel = client
        .channel(`public:campaign_records:${path}:${coordinatorId}_${Math.random().toString(36).substring(2, 7)}`)
        .on('postgres_changes', {
          event: '*',
          schema: 'public',
          table: 'campaign_records',
          filter: `record_type=eq.${path}`
        }, () => {
          firestoreService.getCollectionFiltered<T>(path, coordinatorId).then(callback);
        })
        .subscribe();
    }

    return () => {
      activeSubscribers.delete(subObj);
      if (client && channel) {
        client.removeChannel(channel);
      }
    };
  },

  async getCollectionFiltered<T>(path: string, coordinatorId: string): Promise<T[]> {
    const client = getSupabaseClient();
    const localItems = getLocalList<any>(path);

    if (client) {
      try {
        const { data, error } = await client
          .from('campaign_records')
          .select('record_id, payload, coordinator_id')
          .eq('record_type', path);

        if (!error && data) {
          const supabaseItems = data.map(row => ({
            id: row.record_id,
            coordinator_id: row.coordinator_id,
            ...(row.payload || {})
          })) as any[];

          setLocalList(path, supabaseItems);
          if ((path === 'voters' || path === 'teams') && supabaseItems.length === 0) {
            import('./campaignSeedService').then(m => m.ensureSeedCampaignData()).catch(() => {});
          }

          if (!coordinatorId || coordinatorId === 'all' || coordinatorId === 'geral') return supabaseItems as T[];

          return supabaseItems.filter(item => {
            if (!item) return false;
            const itemCoord = item.coordinatorId || item.coordinator_id;
            const itemReg = item.regionalCoordId;
            const itemRegEmail = item.regionalCoordEmail;
            return (
              itemCoord === coordinatorId ||
              itemCoord === 'geral' ||
              itemReg === coordinatorId ||
              itemRegEmail === coordinatorId ||
              item.coordinatorId === 'geral' ||
              !itemCoord
            );
          }) as T[];
        }
      } catch (e) {
        console.warn(`Supabase getCollectionFiltered error for ${path}:`, e);
      }
    }

    if (!coordinatorId || coordinatorId === 'all' || coordinatorId === 'geral') return localItems as T[];
    return localItems.filter(item => {
      if (!item) return false;
      const itemCoord = item.coordinatorId || item.coordinator_id;
      const itemReg = item.regionalCoordId;
      const itemRegEmail = item.regionalCoordEmail;
      return (
        itemCoord === coordinatorId ||
        itemCoord === 'geral' ||
        itemReg === coordinatorId ||
        itemRegEmail === coordinatorId ||
        item.coordinatorId === 'geral' ||
        !itemCoord
      );
    }) as T[];
  }
};
