import { createClient } from '@supabase/supabase-js';

export function getSupabaseClient() {
  let url = import.meta.env.VITE_SUPABASE_URL || '';
  // user accidentally included /rest/v1/ in env
  if (url.endsWith('/rest/v1/')) {
    url = url.replace('/rest/v1/', '');
  } else if (url.endsWith('/rest/v1')) {
    url = url.replace('/rest/v1', '');
  }
  const key = import.meta.env.VITE_SUPABASE_ANON_KEY || '';
  return createClient(url, key);
}

export const db = {
  async fetchAll() {
    if (!import.meta.env.VITE_SUPABASE_URL) return null;
    const supabase = getSupabaseClient();
    
    // Fetch all tables
    const [
      { data: usersData },
      { data: jobdesksData },
      { data: masterTasksData },
      { data: tasksData },
      { data: logbooksData },
      { data: categoriesData },
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

    // Parse specific structures because some are JSON columns, some are related
    return {
      users: (usersData || []).map((u: any) => ({
        ...u,
        tanggalMulai: u.tanggalmulai,
        tanggalSelesai: u.tanggalselesai,
        nomorSurat: u.nomorsurat || u.nomorSurat || ''
      })),
      jobdesks: (jobdesksData || []).reduce((acc: any, j: Record<string, any>) => ({ ...acc, [j.rolename || j.roleName]: j.description }), {} as Record<string, string>),
      masterTasks: (masterTasksData || []).map((m: any) => ({
        ...m,
        workType: m.worktype,
        targetRole: m.targetrole
      })),
      tasks: (tasksData || []).map((t: any) => ({
        ...t,
        taskId: t.id,
        masterId: t.masterid,
        assignedNim: t.assignednim,
        taskName: t.taskname,
        dateAssigned: t.dateassigned,
        completedDesc: t.completeddesc,
        completedDate: t.completeddate,
        googleDocUrl: t.googledocurl,
        pointsChecked: t.points
      })),
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
      properties: appStateData?.find((s: any) => s.id === 'propertiesData')?.data || null,
      nomorSuratData: appStateData?.find((s: any) => s.id === 'nomorSuratData')?.data || null
    };
  },
  
  // Realtime saving logic
  async runMutation(table: string, method: 'insert' | 'update' | 'delete' | 'upsert', payload: any, matchQuery?: { column: string, value: any }) {
    if (!import.meta.env.VITE_SUPABASE_URL) return;
    const supabase = getSupabaseClient();
    
    let query = method === 'delete' ? supabase.from(table).delete() : supabase.from(table)[method](payload);
    
    if (matchQuery && (method === 'update' || method === 'delete')) {
      query = query.eq(matchQuery.column, matchQuery.value);
    }
    
    const { error } = await query;
    if (error) {
      if (table === 'users' && method === 'upsert' && error.message.includes('nomorsurat')) {
        console.warn("Retrying users upsert without 'nomorsurat' column (schema outdated).");
        const fallbackPayload = Array.isArray(payload) ? payload.map((p: any) => {
          const { nomorsurat, ...rest } = p;
          return rest;
        }) : (() => { const { nomorsurat, ...rest } = payload; return rest; })();
        const fallbackQuery = supabase.from(table)[method](fallbackPayload);
        const { error: fallbackError } = await fallbackQuery;
        if (fallbackError) {
          throw fallbackError;
        }
        return;
      }
      console.error(`Supabase ${method} error on ${table}:`, error.message);
      throw error;
    }
  }
};

