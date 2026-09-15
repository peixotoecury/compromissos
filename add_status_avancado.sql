-- LAWgico Compromissos — evolução pro modelo "Painel de Tarefas" (5 estágios,
-- prioridade, bloqueio, responsável atual). Rodar no SQL Editor do projeto
-- Supabase rpibvjcnrseuugpkfmdj, DEPOIS de add_automacoes.sql.

-- 1) Prioridade por ocorrência
alter table public.compromissos_entregas
  add column if not exists prioridade text not null default 'media'; -- 'alta' | 'media' | 'baixa'

-- 2) Bloqueio (ortogonal ao estágio — uma tarefa "Em andamento" pode estar bloqueada)
alter table public.compromissos_entregas
  add column if not exists bloqueado boolean not null default false;
alter table public.compromissos_entregas
  add column if not exists bloqueado_motivo text;

-- 3) Responsável ATUAL da ocorrência (quando a tarefa passa de mão em mão —
--    ex.: vai pra "Para validação" e passa a precisar de ação de outra
--    pessoa). Null = usa o responsável da definição (comportamento de hoje).
alter table public.compromissos_entregas
  add column if not exists resp_atual_nome text;
alter table public.compromissos_entregas
  add column if not exists resp_atual_email text;

-- 4) Papel/função de quem é responsável (ex.: Advogado, Assistente,
--    Coordenador) — vem da coluna "Função" do Alertas.xlsx (já lida pelo
--    sync_definicoes.py mas não persistida até agora). Usado na tela de
--    Automações pra selecionar destinatários por papel com 1 clique.
alter table public.compromissos_definicoes
  add column if not exists responsavel_papel text;

-- 5) Migra o vocabulário de status de 2 estados pra 5 estágios.
--    'pendente' vira 'a_fazer', 'entregue' vira 'concluido'. Novos valores
--    possíveis a partir de agora: a_fazer | em_andamento | aguardando_terceiro
--    | para_validacao | concluido.
update public.compromissos_entregas set status = 'a_fazer'  where status = 'pendente';
update public.compromissos_entregas set status = 'concluido' where status = 'entregue';

NOTIFY pgrst, 'reload schema';
