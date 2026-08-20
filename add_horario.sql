-- LAWgico Compromissos — adiciona suporte a horário específico do dia
-- Rodar no SQL Editor do projeto Supabase rpibvjcnrseuugpkfmdj (mesmo onde
-- create_tables_compromissos.sql foi rodado).
--
-- Usado pelos alertas curados em Alertas.xlsx (origem='alerta'): cada um
-- tem um horário fixo (ex: '11:00') em que deve ser disparado, diferente do
-- catálogo geral (origem='atividade'/'pauta'/'relatorio'), que não usa esse
-- campo e continua sendo processado só pelo envio único da manhã.

alter table public.compromissos_definicoes
  add column if not exists horario text;  -- formato 'HH:MM', ex: '09:15', '16:00'. Null = sem horário fixo (catálogo geral).

comment on column public.compromissos_definicoes.horario is
  'Horário fixo do alerta (HH:MM), só usado pelos itens origem=alerta (Alertas.xlsx). Null para o catálogo geral.';

NOTIFY pgrst, 'reload schema';
