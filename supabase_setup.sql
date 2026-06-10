-- Execute uma vez no SQL Editor do Supabase.
-- As políticas RLS garantem que cada usuário veja e altere somente os próprios dados.

create table if not exists public.transacoes (
    id uuid primary key default gen_random_uuid(),
    user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
    data date not null,
    descricao text,
    categoria text,
    valor numeric not null,
    tipo text,
    cartao text,
    created_at timestamptz not null default now()
);

create table if not exists public.investimentos (
    id uuid primary key default gen_random_uuid(),
    user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
    data date not null,
    tipo text,
    valor numeric not null,
    rentabilidade text,
    descricao text,
    status text,
    created_at timestamptz not null default now()
);

alter table public.transacoes enable row level security;
alter table public.investimentos enable row level security;

drop policy if exists "usuarios gerenciam as proprias transacoes" on public.transacoes;
create policy "usuarios gerenciam as proprias transacoes"
on public.transacoes
for all
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);

drop policy if exists "usuarios gerenciam os proprios investimentos" on public.investimentos;
create policy "usuarios gerenciam os proprios investimentos"
on public.investimentos
for all
to authenticated
using ((select auth.uid()) = user_id)
with check ((select auth.uid()) = user_id);

grant select, insert, update, delete on public.transacoes to authenticated;
grant select, insert, update, delete on public.investimentos to authenticated;
