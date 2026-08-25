-- LAWgico Compromissos — adiciona justificativa de não entrega
-- Rodar no SQL Editor do projeto Supabase rpibvjcnrseuugpkfmdj (mesmo onde
-- create_tables_compromissos.sql e add_horario.sql foram rodados).
--
-- Permite registrar por que um compromisso pendente/atrasado ainda não foi
-- entregue, sem precisar marcá-lo como entregue pra isso.

alter table public.compromissos_entregas
  add column if not exists justificativa text;

comment on column public.compromissos_entregas.justificativa is
  'Texto livre explicando por que o compromisso ainda não foi entregue. Null = sem justificativa registrada.';

NOTIFY pgrst, 'reload schema';
