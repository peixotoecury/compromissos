-- LAWgico Compromissos — status próprio pras Agendas do dashboard de prazos
-- em tempo real (tabela `prazos_agendas`, mesmo projeto Supabase). NUNCA
-- escreve em `prazos_agendas` — só lê de lá (tipo_agenda='Agenda') e guarda
-- o estágio/prioridade/bloqueio/responsável-atual do Kanban aqui, numa
-- tabela separada, pra não colidir com o sync automático do LM
-- (sync_prazos.py roda a cada 2min e faz upsert+limpeza de obsoletos em
-- prazos_agendas — se o Compromissos escrevesse lá junto, arriscava
-- corromper ou perder a mudança no próximo ciclo).
-- Rodar no SQL Editor do projeto Supabase rpibvjcnrseuugpkfmdj.

create table public.agendas_status_compromissos (
  agenda_id bigint primary key,          -- mesmo id de prazos_agendas (bigint)
  status text not null default 'a_fazer',-- a_fazer | em_andamento | aguardando_terceiro | para_validacao | concluido
  prioridade text not null default 'media',
  bloqueado boolean not null default false,
  bloqueado_motivo text,
  resp_atual_nome text,
  resp_atual_email text,
  atualizado_em timestamptz default now(),
  atualizado_por text
);
alter table public.agendas_status_compromissos enable row level security;
create policy "anon full access" on public.agendas_status_compromissos for all to anon using (true) with check (true);
alter publication supabase_realtime add table public.agendas_status_compromissos;

NOTIFY pgrst, 'reload schema';
