export default {
  async fetch(request, env) {
    const cors = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    if (!env.SUPABASE_URL || !env.SUPABASE_SERVICE_KEY) {
      return new Response(JSON.stringify({ error: 'missing_env' }), {
        status: 500,
        headers: { 'Content-Type': 'application/json', ...cors }
      });
    }

    const endpoint = `${env.SUPABASE_URL}/rest/v1/impressoes`;
    const auth = {
      'apikey': env.SUPABASE_SERVICE_KEY,
      'Authorization': `Bearer ${env.SUPABASE_SERVICE_KEY}`,
    };

    const PAGE = 1000;
    let all = [];
    let offset = 0;

    while (true) {
      const resp = await fetch(
        `${endpoint}?select=dt,mes,pages,impressora,dept,colorida&order=dt.asc,id.asc&limit=${PAGE}&offset=${offset}`,
        { headers: auth }
      );

      if (!resp.ok) {
        return new Response(
          JSON.stringify({ error: 'upstream_error', status: resp.status }),
          { status: 502, headers: { 'Content-Type': 'application/json', ...cors } }
        );
      }

      const batch = await resp.json();
      all = all.concat(batch);
      if (batch.length < PAGE) break;
      offset += PAGE;
    }

    const periodo = buildPeriodo(all);

    return new Response(
      JSON.stringify({ periodo, registros: all.length, data: all }),
      {
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Cache-Control': 'public, max-age=300',
          ...cors
        }
      }
    );
  }
};

function buildPeriodo(data) {
  if (!data.length) return '';
  const dates = data.map(r => r.dt).sort();
  const fmt = s => { const [y, m, d] = s.split('-'); return `${d}/${m}/${y}`; };
  return `${fmt(dates[0])} – ${fmt(dates[dates.length - 1])}`;
}
