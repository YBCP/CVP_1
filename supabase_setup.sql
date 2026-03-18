-- ============================================================
-- CVP - Sistema de Visitas Técnicas
-- Ejecutar en Supabase SQL Editor
-- ============================================================

-- 1. Crear tabla visitas
create table if not exists visitas (
  id bigserial primary key,
  "NUM_VISITA"         text unique not null,
  "FECHA_PROGRAMADA"   text,
  "REA"                text,
  "SIN_REA"            text,
  "DIRECCION_MANUAL"   text,
  "LATITUD_MANUAL"     text,
  "LONGITUD_MANUAL"    text,
  "TECNICOS"           text,
  "ESTADO"             text default 'Pendiente',
  "NUM_VISITA_PREDIO"  text,
  "FECHA_REGISTRO"     text,
  "OBSERVACIONES_PROG" text
);

-- 2. Crear tabla resultados
create table if not exists resultados (
  id bigserial primary key,
  "NUM_VISITA"              text unique not null,
  "REA"                     text,
  "FECHA_VISITA"            text,
  "HORA_INICIO"             text,
  "HORA_FIN"                text,
  "TECNICOS"                text,
  "RESULTADO"               text,
  "OCUPACION"               text,
  "PROP_CONTACTADO"         text,
  "TIPO_CONSTRUCCION"       text,
  "NUM_PISOS"               text,
  "ESTADO_CONSERVACION"     text,
  "LINDERO_NORTE"           text,
  "LINDERO_SUR"             text,
  "LINDERO_ORIENTE"         text,
  "LINDERO_OCCIDENTE"       text,
  "TIPO_INMUEBLE"           text,
  "ESTRATO"                 text,
  "UNIDADES_VIVIENDA"       text,
  "UPL"                     text,
  "UPZ"                     text,
  "AREA_TERRENO"            text,
  "AREA_CONSTRUCCION"       text,
  "TIPO_GESTION"            text,
  "TELEFONO_BENEFICIARIO"   text,
  "CORREO_BENEFICIARIO"     text,
  "COMPONENTE"              text,
  "MOTIVO_FALLIDA"          text,
  "OBSERVACIONES"           text,
  "FOTOS"                   text,
  "FECHA_REGISTRO"          text
);

-- 3. Deshabilitar RLS para acceso con anon key
alter table visitas disable row level security;
alter table resultados disable row level security;

-- 4. (Alternativa segura) Habilitar RLS con políticas permisivas
-- alter table visitas enable row level security;
-- create policy "allow all" on visitas for all using (true) with check (true);
-- alter table resultados enable row level security;
-- create policy "allow all" on resultados for all using (true) with check (true);
