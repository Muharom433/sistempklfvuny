import { createClient } from '@supabase/supabase-js';

export function getSupabaseClient() {
  let url = import.meta.env.VITE_SUPABASE_URL || '';
  // Strip /rest/v1/ suffix if user accidentally included it
  if (url.endsWith('/rest/v1/')) {
    url = url.replace('/rest/v1/', '');
  } else if (url.endsWith('/rest/v1')) {
    url = url.replace('/rest/v1', '');
  }
  // Remove any trailing slashes
  url = url.replace(/\/+$/, '');
  const key = import.meta.env.VITE_SUPABASE_ANON_KEY || '';
  return createClient(url, key);
}

// Map table -> primary key column name used for onConflict in upsert
const UPSERT_CONFLICT_COLUMNS: Record<string, string> = {
  users: 'nim',
  jobdesks: 'rolename',
  master_tasks: 'id',
  tasks: 'id',
  logbooks: 'logbookid',
  categories: 'name',
  app_state: 'id',
};

export const db = {
  async fetchAll() {
    if (!import.meta.env.VITE_SUPABASE_URL) return null;
    const supabase = getSupabaseClient();

    // Fetch all tables in parallel
    const [
      { data: usersData,       error: usersErr },
      { data: jobdesksData },
      { data: masterTasksData, error: masterErr },
      { data: tasksData },
      { data: logbooksData },
      { data: categoriesData,  error: catErr },
      { data: appStateData }
    ] = await Promise.all([
      supabase.from('users').select('*'),
      supabase.from('jobdesks').select('*'),
      supabase.from('master_tasks').select('*'),
      supabase.from('tasks').select('*'),
      supabase.from('logbooks').select('*'),
      supabase.from('categories').select('*'),
      supabase.from('app_state').select('*')
    ]);

    if (catErr)    console.error('[Supabase fetchAll] categories error:', JSON.stringify(catErr));
    if (masterErr) console.error('[Supabase fetchAll] master_tasks error:', JSON.stringify(masterErr));
    if (usersErr)  console.error('[Supabase fetchAll] users error:', JSON.stringify(usersErr));

    return {
      users: (usersData || []).map((u: any) => ({
        ...u,
        tanggalMulai: u.tanggalmulai,
        tanggalSelesai: u.tanggalselesai,
        nomorSurat: u.nomorsurat || u.nomorSurat || ''
      })),
      jobdesks: (jobdesksData || []).reduce(
        (acc: any, j: Record<string, any>) => ({ ...acc, [j.rolename || j.roleName]: j.description }),
        {} as Record<string, string>
      ),
      masterTasks: (masterTasksData || []).map((m: any) => ({
        ...m,
        workType: m.worktype,
        targetRole: m.targetrole
      })),
      tasks: (tasksData || []).map((t: any) => {
        // Parse raw `points` column from Supabase (stored as { checked: bool[], dates: str[] })
        let rawPointsCol = t.points;
        if (typeof rawPointsCol === 'string') {
          try { rawPointsCol = JSON.parse(rawPointsCol); } catch (_) { rawPointsCol = null; }
        }
        const pointsChecked: boolean[] = Array.isArray(rawPointsCol)
          ? rawPointsCol  // legacy: was stored as bool[]
          : (rawPointsCol?.checked || []);
        const checkDates: string[] = Array.isArray(rawPointsCol)
          ? []            // legacy format had no dates
          : (rawPointsCol?.dates || []);

        // Destructure to remove raw `points` from spread so it does NOT shadow
        // the string[] `points` that App.tsx sets after merging with master_tasks.
        const { points: _rawPointsDrop, ...rest } = t;
        return {
          ...rest,
          taskId: t.id,
          masterTaskId: t.masterid,
          assignedNim: t.assignednim,
          taskName: t.taskname,
          dateAssigned: t.dateassigned,
          completedDesc: t.completeddesc,
          completedDate: t.completeddate,
          googleDocUrl: t.googledocurl,
          // ── Tiga integer sumber kebenaran persentase ──
          pointsTotal:    typeof t.pointstotal   === 'number' ? t.pointstotal   : null,
          dbCheckedCount: typeof t.checkedcount  === 'number' ? t.checkedcount  : null,
          percentage:     typeof t.percentage    === 'number' ? t.percentage    : null,
          // `points` (string[]) will be filled by App.tsx repair logic from master_tasks
          points: undefined,
          pointsChecked,
          checkDates,
        };
      }),
      logbooks: (logbooksData || []).map((l: any) => ({
        ...l,
        logbookId: l.logbookid,
        taskId: l.taskid,
        taskName: l.taskname,
        timestamp: l.date,
        workDescription: l.workdescription,
        hoursSpent: l.hoursspent,
        notes: l.gradenote,
        googleDocUrl: l.googledocurl
      })),
      categories: (categoriesData || []).map((c: any) => c.name).filter(Boolean),
      categoriesData: (() => {
        const catState = appStateData?.find((s: any) => s.id === 'categoriesData');
        if (!catState) return null;
        const d = catState.data;
        return Array.isArray(d) ? d : (typeof d === 'string' ? JSON.parse(d) : null);
      })(),
      properties: appStateData?.find((s: any) => s.id === 'propertiesData')?.data || null,
      nomorSuratData: appStateData?.find((s: any) => s.id === 'nomorSuratData')?.data || null
    };
  },

  /**
   * Run a single mutation against a Supabase table.
   * For 'upsert', automatically uses the correct conflict column so duplicates are
   * updated rather than rejected.
   */
  async runMutation(
    table: string,
    method: 'insert' | 'update' | 'delete' | 'upsert',
    payload: any,
    matchQuery?: { column: string; value: any }
  ) {
    if (!import.meta.env.VITE_SUPABASE_URL) return;
    const supabase = getSupabaseClient();

    let query: any;
    if (method === 'delete') {
      query = supabase.from(table).delete();
    } else if (method === 'upsert') {
      const conflictCol = UPSERT_CONFLICT_COLUMNS[table];
      query = conflictCol
        ? supabase.from(table).upsert(payload, { onConflict: conflictCol })
        : supabase.from(table).upsert(payload);
    } else {
      query = supabase.from(table)[method](payload);
    }

    if (matchQuery && (method === 'update' || method === 'delete')) {
      query = query.eq(matchQuery.column, matchQuery.value);
    }

    const { error } = await query;

    if (error) {
      // Special retry for users table when nomorsurat column doesn't exist yet
      if (table === 'users' && method === 'upsert' && error.message?.includes('nomorsurat')) {
        console.warn("[Supabase] Retrying users upsert without 'nomorsurat' column (schema outdated).");
        const fallbackPayload = Array.isArray(payload)
          ? payload.map((p: any) => { const { nomorsurat, ...rest } = p; return rest; })
          : (() => { const { nomorsurat, ...rest } = payload; return rest; })();
        const conflictCol = UPSERT_CONFLICT_COLUMNS[table];
        const fallbackQuery = conflictCol
          ? supabase.from(table).upsert(fallbackPayload, { onConflict: conflictCol })
          : supabase.from(table).upsert(fallbackPayload);
        const { error: fallbackError } = await fallbackQuery;
        if (fallbackError) {
          console.error(`[Supabase] Fallback upsert error on ${table}:`, JSON.stringify(fallbackError));
          throw fallbackError;
        }
        return;
      }
      console.error(
        `[Supabase] ${method} error on table="${table}":`,
        JSON.stringify(error),
        '| payload sample:',
        JSON.stringify(Array.isArray(payload) ? payload[0] : payload)
      );
      // Graceful retry for tasks: jika kolom baru belum ada di schema, coba tanpa mereka
      if (table === 'tasks' && method === 'upsert' &&
          (error.message?.includes('pointstotal') || error.message?.includes('checkedcount') || error.message?.includes('percentage'))) {
        console.warn('[Supabase] Retrying tasks upsert without new integer columns (schema belum di-ALTER).');
        const strip = (p: any) => { const { pointstotal, checkedcount, percentage, ...r } = p; return r; };
        const fallback = Array.isArray(payload) ? payload.map(strip) : strip(payload);
        const conflictCol = UPSERT_CONFLICT_COLUMNS[table];
        const fbQuery = conflictCol
          ? supabase.from(table).upsert(fallback, { onConflict: conflictCol })
          : supabase.from(table).upsert(fallback);
        const { error: fbErr } = await fbQuery;
        if (fbErr) console.error('[Supabase] Fallback tasks upsert error:', JSON.stringify(fbErr));
        return;
      }
      throw error;
    }

    console.log(`[Supabase] ${method} on "${table}" OK`, Array.isArray(payload) ? `(${payload.length} rows)` : '');
  }
};
