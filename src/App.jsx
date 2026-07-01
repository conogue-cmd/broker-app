import { useState, useEffect, useRef } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

// ─── DATOS INICIALES (se migran a Supabase la primera vez que se abre la app) ─
const PROPIEDADES_INICIALES = [
  {"unidad":"1","titular":"GC","valor_compra":"90700","valor_mercado":"90000","direccion":"Alvarez Thomas 1189 3ºA - Cap.Fed","ambientes":"2","baulera":"","cochera":"","mts2":"58","alq_1":"250000","alq_2":"500000","fecha_pago":"1 a 10 de cada mes","expensas":"35000","estado":"alquilado","inquilino":"FEDERICO DICK","dni":"27235996","periodo":"14/02/2024 A 13/02/2026","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"15-6707-2856","mail":"federicodick@yahoo.com.ar","deposito":"","impuesto":"PAGA INQUILINA","nro_partida":"3900290 DV04 Y 3900305 DV01","admin_edificio":"A.Thomas Loredana tel.3969-6395"},
  {"unidad":"2","titular":"EC","valor_compra":"63500","valor_mercado":"","direccion":"Alvarez Thomas 1189 3ºB - Cap.Fed.","ambientes":"1 y medio","baulera":"","cochera":"","mts2":"38","alq_1":"550000","alq_2":"600000","fecha_pago":"1 a 10 de cada mes","expensas":"20000","estado":"alquilado","inquilino":"ALAN GABRIEL NUZZOLESE","dni":"29193568","periodo":"01/10/2025 A 30/09/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"15-2832-8464","mail":"elsitiodesul@yahoo.com.ar","deposito":"550","impuesto":"PAGA INQUILINO","nro_partida":"3900291 DV00","admin_edificio":"A.Thomas Loredana tel.3969-6395"},
  {"unidad":"3","titular":"EC","valor_compra":"82500","valor_mercado":"85000","direccion":"Belgrano 631 UF 8 - San Fernando","ambientes":"4","baulera":"","cochera":"SI","mts2":"106","alq_1":"13900","alq_2":"21500","fecha_pago":"1 a 10 de cada mes","expensas":"3500","estado":"alquilado","inquilino":"DANIELA LANFRANCO","dni":"26122319","periodo":"27/05/2020 A 26/05/2022","inmobiliaria":"","comision":"0","telefono":"","mail":"danielalanfranco@hotmail.com","deposito":"12586","impuesto":"","nro_partida":"096-050011-3","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"4","titular":"GC","valor_compra":"79000","valor_mercado":"","direccion":"Alvarez Thomas 1189 6ºB - Cap.Fed.","ambientes":"1 y medio","baulera":"","cochera":"","mts2":"38","alq_1":"130000","alq_2":"260000","fecha_pago":"1 a 10 de cada mes","expensas":"1500","estado":"alquilado","inquilino":"REBOLLEDO NICOLAS JULIO","dni":"37969146","periodo":"26/08/2025 A 31/08/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"rebolledo.nr@gmail.com","deposito":"550","impuesto":"PAGAMOS NOSOTROS","nro_partida":"3900297 DV09","admin_edificio":"A.Thomas Loredana tel.3969-6395"},
  {"unidad":"5","titular":"CC","valor_compra":"89500","valor_mercado":"","direccion":"Belgrano 1132 depto 2B - San Fernando","ambientes":"2","baulera":"","cochera":"SI","mts2":"54","alq_1":"85000","alq_2":"160000","fecha_pago":"1 a 10 de cada mes","expensas":"4500","estado":"alquilado","inquilino":"CONSTANZA INES CANGIALOSI","dni":"34271864","periodo":"01/07/2023 A 30/06/2025","inmobiliaria":"SOLODUEÑOS","comision":"0","telefono":"","mail":"","deposito":"$6000 + U$S 67","impuesto":"","nro_partida":"096-53034 Y 096-53026","admin_edificio":"consorciobelgrano1132@gmail.com"},
  {"unidad":"6","titular":"EC","valor_compra":"89500","valor_mercado":"90000","direccion":"Belgrano 1132 depto 4B - San Fernando","ambientes":"3","baulera":"","cochera":"SI","mts2":"54","alq_1":"48000","alq_2":"72000","fecha_pago":"1 a 10 de cada mes","expensas":"3500","estado":"libre","inquilino":"","dni":"","periodo":"","inmobiliaria":"","comision":"","telefono":"1535049195","mail":"adriantrompeta65@gmail.com","deposito":"1848","impuesto":"","nro_partida":"096-53038 Y 096-53030","admin_edificio":"consorciobelgrano1132@gmail.com"},
  {"unidad":"7","titular":"CC","valor_compra":"26000","valor_mercado":"31500","direccion":"Tres Arroyos 2771 - 1C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"18.5","alq_1":"40000","alq_2":"70000","fecha_pago":"1 a 10 de cada mes","expensas":"6000","estado":"libre","inquilino":"CARLOS JOSE AGUILERA CORREA","dni":"96204064","periodo":"24/11/2022 A 23/11/2025","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"aguilera.carlos.2000@gmail.com","deposito":"-118853","impuesto":"PAGAMOS NOSOTROS","nro_partida":"3949570-00","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"8","titular":"GC","valor_compra":"","valor_mercado":"U$S 48000","direccion":"Tres Arroyos 2771 - 3C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"18.5","alq_1":"","alq_2":"","fecha_pago":"","expensas":"200","estado":"vendido","inquilino":"","dni":"","periodo":"","inmobiliaria":"PIGNATARO","comision":"","telefono":"","mail":"","deposito":"","impuesto":"","nro_partida":"3949577-05","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"9","titular":"EC","valor_compra":"","valor_mercado":"U$S 66000","direccion":"Tres Arroyos 2771 - 3A - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"30.92","alq_1":"","alq_2":"","fecha_pago":"","expensas":"1000","estado":"vendido","inquilino":"","dni":"","periodo":"","inmobiliaria":"PIGNATARO","comision":"0.02","telefono":"","mail":"","deposito":"","impuesto":"","nro_partida":"3949575-02","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"10","titular":"GC","valor_compra":"","valor_mercado":"","direccion":"Tres Arroyos 2771 - 4B - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"33.25","alq_1":"","alq_2":"","fecha_pago":"","expensas":"300","estado":"vendido","inquilino":"","dni":"","periodo":"30/06/2017 A 30/09/2018","inmobiliaria":"SOLODUEÑOS","comision":"","telefono":"","mail":"","deposito":"0","impuesto":"","nro_partida":"3949580-06","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"11","titular":"CC","valor_compra":"","valor_mercado":"U$S 89000","direccion":"Tres Arroyos 2771 - 5C - CABA","ambientes":"1 Y 1/2","baulera":"SI NRO.1","cochera":"SI NRO.3","mts2":"45","alq_1":"","alq_2":"","fecha_pago":"","expensas":"600","estado":"vendido","inquilino":"","dni":"","periodo":"VENDIDO EL 09/08/2018","inmobiliaria":"PIGNATARO","comision":"0.02","telefono":"","mail":"","deposito":"0","impuesto":"","nro_partida":"3949584-01","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"12","titular":"GC","valor_compra":"50500","valor_mercado":"","direccion":"Tres Arroyos 2771 - 6C - CABA","ambientes":"2","baulera":"SI NRO.7","cochera":"SI NRO. 4","mts2":"36.27","alq_1":"65000","alq_2":"104000","fecha_pago":"1 a 10 de cada mes","expensas":"6000","estado":"alquilado","inquilino":"CAMILA CARESANI","dni":"41104564","periodo":"01/09/2022 A 31/08/2025","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"camilacaresani@gmail.com","deposito":"-17226","impuesto":"PAGAMOS NOSOTROS","nro_partida":"3949587-00","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"13","titular":"GC","valor_compra":"100000","valor_mercado":"","direccion":"Lavalle 167 - 2D - San Fernando","ambientes":"2","baulera":"","cochera":"SI NRO. 12","mts2":"45","alq_1":"70000","alq_2":"140000","fecha_pago":"1 a 10 de cada mes","expensas":"8000","estado":"alquilado","inquilino":"LUNA BELEN RODRIGUEZ MAROTTA","dni":"42374900","periodo":"01/03/2023 A 28/02/2026","inmobiliaria":"LVC PROPIEDADES","comision":"0.05","telefono":"+54911-5479-0664","mail":"","deposito":"","impuesto":"","nro_partida":"096-056168-6","admin_edificio":"ADM Lavalle Veronica Criado +54 9 11 5997-4729"},
  {"unidad":"14","titular":"CC","valor_compra":"51000","valor_mercado":"45000","direccion":"Artigas 333 - 3A - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"31","alq_1":"45000","alq_2":"70000","fecha_pago":"1 a 10 de cada mes","expensas":"800","estado":"alquilado","inquilino":"","dni":"44790327","periodo":"01/09/2022 A 31/08/2025","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"","deposito":"0.45","impuesto":"","nro_partida":"4137697","admin_edificio":"info@pehache.com.ar"},
  {"unidad":"15","titular":"CC","valor_compra":"51000","valor_mercado":"","direccion":"Artigas 333 - 3B - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"31","alq_1":"48000","alq_2":"80000","fecha_pago":"1 a 10 de cada mes","expensas":"1500","estado":"alquilado","inquilino":"BENITO KIM","dni":"33153823","periodo":"12/10/2025 A 11/10/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"benitokim87@hotmail.com","deposito":"400","impuesto":"","nro_partida":"4137699","admin_edificio":"info@pehache.com.ar"},
  {"unidad":"16","titular":"EC","valor_compra":"51000","valor_mercado":"","direccion":"Artigas 333 - 3C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"31","alq_1":"33000","alq_2":"50000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"YENY LORENA VELASQUEZ PULIDO","dni":"94488242","periodo":"08/05/2025 A 07/05/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"lors_20@hotmail.com","deposito":"-9616","impuesto":"","nro_partida":"4137698","admin_edificio":"info@pehache.com.ar"},
  {"unidad":"17","titular":"EC","valor_compra":"51000","valor_mercado":"","direccion":"Artigas 333 - 6B - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"31","alq_1":"22000","alq_2":"32000","fecha_pago":"1 a 10 de cada mes","expensas":"3500","estado":"alquilado","inquilino":"JOSEPH BRAYAN HUANCA CONDORI","dni":"94345301","periodo":"07/10/2024 A 06/10/2026","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"joseph.daykel107@gmail.com","deposito":"U$S 300","impuesto":"","nro_partida":"4137712","admin_edificio":"info@pehache.com.ar"},
  {"unidad":"18","titular":"GC","valor_compra":"65000","valor_mercado":"","direccion":"Gurruchaga 381 - 2B - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"35.69","alq_1":"400000","alq_2":"450000","fecha_pago":"1 a 10 de cada mes","expensas":"100000","estado":"alquilado","inquilino":"NICOLAS IGNACIO RUSSOMANNO","dni":"40511858","periodo":"21/04/2025 A 20/03/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"115801-8531","mail":"russomannonicolas@gmail.com","deposito":"U$S 400","impuesto":"","nro_partida":"4127814-06","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"19","titular":"CC","valor_compra":"41500","valor_mercado":"48500","direccion":"Gurruchaga 381 - 2C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"22.79","alq_1":"22000","alq_2":"31900","fecha_pago":"1 a 10 de cada mes","expensas":"3500","estado":"alquilado","inquilino":"JULIO CESAR BLANCO AVILAN","dni":"95799401","periodo":"12/08/2021 A 11/08/2024","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"juliob412@gmail.com","deposito":"22000","impuesto":"","nro_partida":"4127816-09","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"20","titular":"EC","valor_compra":"65000","valor_mercado":"","direccion":"Gurruchaga 381 - 3B - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"36.55","alq_1":"400000","alq_2":"450000","fecha_pago":"1 a 10 de cada mes","expensas":"100000","estado":"alquilado","inquilino":"GUADALUPE DANIELA MARTI","dni":"42093962","periodo":"17/04/2026 A 16/04/2028","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"marti.guada25@gmail.com","deposito":"U$S 500","impuesto":"","nro_partida":"4127818-01","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"21","titular":"GC","valor_compra":"72000","valor_mercado":"71000","direccion":"Gurruchaga 381 - 3D - CABA","ambientes":"1","baulera":"","cochera":"SI","mts2":"39.77","alq_1":"36000","alq_2":"54000","fecha_pago":"1 a 10 de cada mes","expensas":"6500","estado":"alquilado","inquilino":"IGNACION NOLTE POLLEDO","dni":"34123716","periodo":"01/04/2024 A 31/03/2026","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"156532-5922","mail":"nachonolte@hotmail.com","deposito":"0","impuesto":"","nro_partida":"4127819-08","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"22","titular":"EC","valor_compra":"76000","valor_mercado":"","direccion":"Gurruchaga 381 - 4A - CABA","ambientes":"1","baulera":"","cochera":"SI","mts2":"42.72","alq_1":"400000","alq_2":"450000","fecha_pago":"1 a 10 de cada mes","expensas":"100000","estado":"alquilado","inquilino":"NADIA BARRIOS","dni":"40007377","periodo":"06/06/2025 A 05/06/2027","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"nadia.bbarrios@gmail.com","deposito":"U$S 400","impuesto":"","nro_partida":"4127821-02","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"23","titular":"CC","valor_compra":"70000","valor_mercado":"63000","direccion":"Gurruchaga 381 - 5D - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"39.77","alq_1":"20000","alq_2":"27000","fecha_pago":"1 a 10 de cada mes","expensas":"3500","estado":"alquilado","inquilino":"ANDREA CAROLINA ARELLANO SANCHEZ","dni":"95.876.353","periodo":"18/11/2020 A 17/11/2023","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"andreakrolina@gmail.com","deposito":"15495","impuesto":"","nro_partida":"4127827-00","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"24","titular":"EC","valor_compra":"41500","valor_mercado":"45000","direccion":"Gurruchaga 381 - 6C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"22.79","alq_1":"13000","alq_2":"18000","fecha_pago":"1 a 10 de cada mes","expensas":"4500","estado":"alquilado","inquilino":"ADRIANA YICEL CAVALHEIRO","dni":"34892336","periodo":"15/09/2020 A 14/09/2023","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"yicelaborges@hotmail.com","deposito":"0","impuesto":"","nro_partida":"4127832-04","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"25","titular":"EC","valor_compra":"72000","valor_mercado":"68000","direccion":"Gurruchaga 381 - 7D - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"39.77","alq_1":"50000","alq_2":"80000","fecha_pago":"1 a 10 de cada mes","expensas":"8000","estado":"alquilado","inquilino":"RENZO GABRIEL ARTEAGA SANZ","dni":"95979108","periodo":"03/10/2022 A 03/10/2025","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"renzogabriel05@hotmail.com","deposito":"-72977","impuesto":"","nro_partida":"4127835-03","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"26","titular":"CC","valor_compra":"74000","valor_mercado":"70000","direccion":"Gurruchaga 381 - 8A - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"42.72","alq_1":"45000","alq_2":"76500","fecha_pago":"1 a 10 de cada mes","expensas":"9000","estado":"alquilado","inquilino":"LUDMILA VIANA AMIN","dni":"33868326","periodo":"11/10/2022 A 10/10/2024","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"","deposito":"38170","impuesto":"","nro_partida":"4127837-06","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"27","titular":"GC","valor_compra":"45000","valor_mercado":"56000","direccion":"Gurruchaga 381 - 8C - CABA","ambientes":"1","baulera":"","cochera":"","mts2":"25.04","alq_1":"25000","alq_2":"37500","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"","dni":"","periodo":"31/01/2022 A 30/01/2024","inmobiliaria":"PIGNATARO","comision":"0.05","telefono":"","mail":"","deposito":"","impuesto":"","nro_partida":"4127840-07","admin_edificio":"A.Thomas Guadalupe tel.3969-6395"},
  {"unidad":"28","titular":"CC","valor_compra":"145000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 2A - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 5","mts2":"84.31","alq_1":"75000","alq_2":"130000","fecha_pago":"1 a 10 de cada mes","expensas":"4500","estado":"alquilado","inquilino":"MELINA PILAR CARRARO","dni":"42399378","periodo":"01/02/2025 A 31/01/2027","inmobiliaria":"DE LEO PROPIEDADES","comision":"0.05","telefono":"","mail":"","deposito":"U$S 500","impuesto":"","nro_partida":"057-137833-7","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"29","titular":"EC","valor_compra":"157000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 3B - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 23","mts2":"95","alq_1":"86000","alq_2":"155000","fecha_pago":"1 a 10 de cada mes","expensas":"15000","estado":"alquilado","inquilino":"LUIS ALBERTO KEUROGLIAN SAGHATIAN","dni":"30-71614300-3","periodo":"01/03/2025 A 28/02/2027","inmobiliaria":"SOLODUEÑOS","comision":"0","telefono":"","mail":"luis@arquifoam.com","deposito":"79949","impuesto":"","nro_partida":"057-137843-4","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"30","titular":"GC","valor_compra":"167000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 3D/3F - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 25","mts2":"95.5","alq_1":"250000","alq_2":"400000","fecha_pago":"1 a 10 de cada mes","expensas":"","estado":"alquilado","inquilino":"ANA LAURA GOYECHEA","dni":"35247714","periodo":"01/01/2026 A 31/12/2027","inmobiliaria":"SOLODUEÑOS","comision":"0","telefono":"","mail":"ana_goye@hotmail.com","deposito":"U$S 250","impuesto":"","nro_partida":"057-137845-0","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"31","titular":"GC","valor_compra":"136000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 4F - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 40","mts2":"83.8","alq_1":"53000","alq_2":"76000","fecha_pago":"1 a 10 de cada mes","expensas":"9000","estado":"alquilado","inquilino":"TAIPE VILCA CARLOS MANUEL","dni":"19080067","periodo":"16/02/2022 A 15/02/2025","inmobiliaria":"REMAX","comision":"0","telefono":"+54 9 3413 75-6455","mail":"manuelsunrise@gmail.com","deposito":"553000","impuesto":"","nro_partida":"057-137856-6","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"32","titular":"EC","valor_compra":"195000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 5A - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 38","mts2":"114.2","alq_1":"15000","alq_2":"24000","fecha_pago":"1 a 10 de cada mes","expensas":"","estado":"alquilado","inquilino":"CECILIA CASTILLO","dni":"23668321","periodo":"","inmobiliaria":"","comision":"0","telefono":"","mail":"","deposito":"","impuesto":"","nro_partida":"057-137860-4","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"33","titular":"CC","valor_compra":"194000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 5B - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 26","mts2":"113.5","alq_1":"260000","alq_2":"450000","fecha_pago":"1 a 10 de cada mes","expensas":"70000","estado":"alquilado","inquilino":"JORGE MIGUEL MISU","dni":"18218245","periodo":"01/03/2026 A 29/02/2028","inmobiliaria":"LVC PROPIEDADES","comision":"0.05","telefono":"3548-8216","mail":"jmmisuconstrucciones@gmail.com","deposito":"$260000 + USD 215","impuesto":"","nro_partida":"057-137861-2","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"34","titular":"GC","valor_compra":"185000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 5E - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 11","mts2":"109.2","alq_1":"550000","alq_2":"650000","fecha_pago":"1 a 10 de cada mes","expensas":"2500","estado":"alquilado","inquilino":"FACUNDO FERMANI","dni":"42823913","periodo":"01/03/2025 A 28/02/2027","inmobiliaria":"PAOLA DE LEO","comision":"0.05","telefono":"+54 9 11 3610-0025","mail":"fermafacu@gmail.com","deposito":"U$S 500","impuesto":"","nro_partida":"057-137864-7","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"35","titular":"CC","valor_compra":"185000","valor_mercado":"","direccion":"Boulogne Sur Mer 486 - 5F - Pacheco","ambientes":"2","baulera":"SI","cochera":"SI NRO. 6","mts2":"109.2","alq_1":"27000","alq_2":"35000","fecha_pago":"1 a 10 de cada mes","expensas":"","estado":"alquilado","inquilino":"DARIO SENISE","dni":"26312886","periodo":"01/11/2024 A 31/10/2026","inmobiliaria":"SOLODUEÑOS","comision":"0","telefono":"1566478110","mail":"senise11@gmail.com","deposito":"410","impuesto":"","nro_partida":"057-137865-5","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"36","titular":"CC","valor_compra":"110000","valor_mercado":"120000","direccion":"Cabildo 672 - 1B - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 2","mts2":"56.55","alq_1":"550000","alq_2":"650000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"LUCIANO CARLOS PEÑA","dni":"25390686","periodo":"15/03/2025 A 28/02/2027","inmobiliaria":"PAOLA DE LEO","comision":"0.05","telefono":"","mail":"lucianopenia1976@gmail.com","deposito":"U$S 500","impuesto":"","nro_partida":"057-139756","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"37","titular":"GC","valor_compra":"127000","valor_mercado":"","direccion":"Cabildo 672 - 2A - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 3","mts2":"66.25","alq_1":"29000","alq_2":"35000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"FRANCA GALFRASCOLI","dni":"35272351","periodo":"01/11/2023 A 31/10/2026","inmobiliaria":"ZONAPROP","comision":"0","telefono":"1166280667","mail":"galfrascolifranca@gmail.com","deposito":"u$s 3000","impuesto":"","nro_partida":"057-139757","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"38","titular":"EC","valor_compra":"114000","valor_mercado":"","direccion":"Cabildo 672 - 2D - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 6","mts2":"59.55","alq_1":"60000","alq_2":"120000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"ORIANA AGOSTINA LUCIA","dni":"38069455","periodo":"04/07/2025 A 03/07/2027","inmobiliaria":"","comision":"0","telefono":"+54911 6537 8755","mail":"lasluchias@hotmail.com","deposito":"U$S 2400","impuesto":"","nro_partida":"057-139760","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"39","titular":"EC","valor_compra":"110000","valor_mercado":"","direccion":"Cabildo 672 - 2F - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO 8","mts2":"56.55","alq_1":"400000","alq_2":"500000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"VANESA NATALIA FERRERO Y MARCOS JAVIER MUZI","dni":"30.365.328 / 29.732.532","periodo":"01/07/2024 A 30/06/2026","inmobiliaria":"PAOLA DE LEO","comision":"0.05","telefono":"54 9 11 5471-9728","mail":"vafer32@gmail.com","deposito":"U$S 400","impuesto":"","nro_partida":"057-139762","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"40","titular":"EC","valor_compra":"127000","valor_mercado":"","direccion":"Cabildo 672 - 3A - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NR. 13","mts2":"66.25","alq_1":"420000","alq_2":"600000","fecha_pago":"1 a 10 de cada mes","expensas":"80000","estado":"alquilado","inquilino":"YAMILA LORENA RUIZ","dni":"34730950","periodo":"01/09/2024 A 31/08/2026","inmobiliaria":"SOLODUEÑOS","comision":"0","telefono":"+54911 6979 9070","mail":"yamilalruiz@hotmail.com","deposito":"U$S 300","impuesto":"","nro_partida":"057-139763","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"41","titular":"GC","valor_compra":"110000","valor_mercado":"","direccion":"Cabildo 672 - 3F - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 19","mts2":"56.55","alq_1":"70000","alq_2":"112000","fecha_pago":"1 a 10 de cada mes","expensas":"8000","estado":"alquilado","inquilino":"SEBASTIAN IGLESIA","dni":"36475258","periodo":"01/09/2025 A 31/08/2027","inmobiliaria":"ZONAPROP","comision":"0","telefono":"6010-6723","mail":"sebass_iglesia@hotmail.com","deposito":"U$S 1200","impuesto":"","nro_partida":"057-139768","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"42","titular":"GC","valor_compra":"110000","valor_mercado":"","direccion":"Cabildo 672 - 4B - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 12","mts2":"55.85","alq_1":"27000","alq_2":"35000","fecha_pago":"1 a 10 de cada mes","expensas":"4500","estado":"alquilado","inquilino":"ADRIAN CESAR MARCHINO / LINDA URRELS","dni":"38399175","periodo":"01/07/2024 A 30/06/2026","inmobiliaria":"ZONAPROP","comision":"0","telefono":"1144007530","mail":"adrian_marchino@hotmail.com","deposito":"$29000 + U$S 270","impuesto":"","nro_partida":"057-139770","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"43","titular":"CC","valor_compra":"115000","valor_mercado":"","direccion":"Cabildo 672 - 4D - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 9","mts2":"59.55","alq_1":"420000","alq_2":"800000","fecha_pago":"1 a 10 de cada mes","expensas":"5000","estado":"alquilado","inquilino":"LUCIANO CARLOS PEÑA","dni":"25390686","periodo":"15/03/2025 A 28/02/2027","inmobiliaria":"PAOLA DE LEO","comision":"0.05","telefono":"","mail":"lucianopenia1976@gmail.com","deposito":"U$S 500","impuesto":"","nro_partida":"057-139772","admin_edificio":"Administracion C&S 155976-3830"},
  {"unidad":"44","titular":"GC","valor_compra":"90000","valor_mercado":"USD 108000","direccion":"Rivadavia 683 - 1E - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 9","mts2":"55.86","alq_1":"65000","alq_2":"104000","fecha_pago":"1 a 10 de cada mes","expensas":"5500","estado":"alquilado","inquilino":"MANCINELLI MATIAS","dni":"38920387","periodo":"10/08/2022 A 09/08/2025","inmobiliaria":"REMAX","comision":"0","telefono":"","mail":"celexpressargentina@gmail.com","deposito":"93725","impuesto":"","nro_partida":"057-141952-1","admin_edificio":"De Leo propiedades"},
  {"unidad":"45","titular":"EC","valor_compra":"90000","valor_mercado":"","direccion":"Rivadavia 683 - 2E - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 10","mts2":"55.86","alq_1":"33000","alq_2":"43000","fecha_pago":"1 a 10 de cada mes","expensas":"5500","estado":"alquilado","inquilino":"ALEJANDRO SOLGA","dni":"37596361","periodo":"01/02/2026 A 31/01/2028","inmobiliaria":"DIRECTO","comision":"0","telefono":"+54 9 1156165240","mail":"alejandro_solga@hotmail.com","deposito":"$33000 + U$S 220","impuesto":"","nro_partida":"057-141958-0","admin_edificio":"De Leo propiedades"},
  {"unidad":"46","titular":"CC","valor_compra":"90000","valor_mercado":"","direccion":"Rivadavia 683 - 2B - Pacheco","ambientes":"2","baulera":"NO","cochera":"SI NRO. 8","mts2":"55.86","alq_1":"420000","alq_2":"500000","fecha_pago":"1 a 10 de cada mes","expensas":"80000","estado":"alquilado","inquilino":"MARIA CRISTINA ALVAREZ","dni":"5713980","periodo":"01/07/2024 A 30/06/2026","inmobiliaria":"CABRERA","comision":"0","telefono":"15-6529-1424","mail":"nico2205@hotmail.com","deposito":"$33000 + U$S 300","impuesto":"","nro_partida":"057-141955-6","admin_edificio":"De Leo propiedades"}
];

// ─── UTILIDADES ───────────────────────────────────────────────────────────────
function parseFecha(str) {
  if (!str) return null;
  const clean = str.trim().toUpperCase();
  let m = clean.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    let y = parseInt(m[3]); if (y < 100) y += 2000;
    return new Date(y, parseInt(m[2]) - 1, parseInt(m[1]));
  }
  return null;
}

function vencimientoContrato(periodo) {
  if (!periodo) return null;
  const parts = periodo.split(/\s+A\s+/i);
  if (parts.length < 2) return null;
  return parseFecha(parts[1].trim());
}

function diasParaVencer(periodo) {
  const fv = vencimientoContrato(periodo);
  if (!fv) return null;
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  return Math.round((fv - hoy) / 86400000);
}

function semaforo(prop) {
  if (prop.estado === "vendido") return "vendido";
  if (prop.estado === "libre") return "libre";
  const dias = diasParaVencer(prop.periodo);
  if (dias === null) return "sin-fecha";
  if (dias < 0) return "vencido";
  if (dias <= 60) return "proximo";
  return "vigente";
}

function fmtARS(n) {
  if (!n && n !== 0) return "—";
  return "$" + Number(n).toLocaleString("es-AR");
}

function fmtFecha(d) {
  if (!d) return "—";
  return d.toLocaleDateString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric" });
}

const COLORES_SEMAFORO = {
  vigente: { bg: "#eaf3de", color: "#a6e3a1", label: "Vigente" },
  proximo: { bg: "#faeeda", color: "#fab387", label: "Próximo a vencer" },
  vencido: { bg: "#fcebeb", color: "#f38ba8", label: "Vencido" },
  libre: { bg: "#e6f1fb", color: "#89b4fa", label: "Libre / Sin inquilino" },
  vendido: { bg: "#f1efe8", color: "#c0c0e0", label: "Vendido" },
  "sin-fecha": { bg: "#fbeaf0", color: "#cba6f7", label: "Sin fecha" },
};

// ─── COMPONENTES VISUALES ──────────────────────────────────────────────────
function Badge({ tipo }) {
  const c = COLORES_SEMAFORO[tipo] || COLORES_SEMAFORO.libre;
  return (
    <span style={{ background: c.bg, color: c.color, borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 500, whiteSpace: "nowrap" }}>
      {c.label}
    </span>
  );
}

function KPI({ label, value, sub, color }) {
  return (
    <div style={{ background: "#252538", borderRadius: 8, padding: "14px 18px", minWidth: 110 }}>
      <div style={{ fontSize: 12, color: "#c0c0e0", marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 600, color: color || "#e0e0f0" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: "#c0c0e0", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

// ─── MODAL DETALLE ────────────────────────────────────────────────────────────
function ModalDetalle({ prop, onClose, onSave, onDelete }) {
  const [editando, setEditando] = useState(false);
  const [confirmandoBorrado, setConfirmandoBorrado] = useState(false);
  const [form, setForm] = useState({ ...prop });

  const campo = (label, key, tipo = "text") => (
    <div style={{ marginBottom: 10 }}>
      <div style={{ fontSize: 11, color: "#c0c0e0", marginBottom: 2 }}>{label}</div>
      {editando ? (
        <input
          type={tipo}
          value={form[key] || ""}
          onChange={(e) => setForm((f) => ({ ...f, [key]: e.target.value }))}
          style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, background: "#252538" }}
        />
      ) : (
        <div style={{ fontSize: 13, color: "#e0e0f0", padding: "4px 0", borderBottom: "0.5px solid #eee" }}>
          {prop[key] || <span style={{ color: "#4a4a6a" }}>—</span>}
        </div>
      )}
    </div>
  );

  const dias = diasParaVencer(prop.periodo);
  const sem = semaforo(prop);

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#252538", borderRadius: 16, width: "min(700px, 96vw)", maxHeight: "90vh", overflow: "auto", padding: 28, boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, color: "#e0e0f0", marginBottom: 6 }}>{prop.direccion}</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <Badge tipo={sem} />
              {dias !== null && sem !== "vendido" && (
                <span style={{ fontSize: 11, color: "#c0c0e0", padding: "2px 8px", background: "#2d2b55", borderRadius: 4 }}>
                  {dias < 0 ? `Venció hace ${Math.abs(dias)} días` : `Vence en ${dias} días`}
                </span>
              )}
              <span style={{ fontSize: 11, color: "#c0c0e0", padding: "2px 8px", background: "#2d2b55", borderRadius: 4 }}>Titular: {prop.titular}</span>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {editando ? (
              <>
                <button onClick={() => { onSave(form); setEditando(false); }} style={{ padding: "6px 14px", background: "#7c6fcd", color: "#e0e0f0", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Guardar</button>
                <button onClick={() => { setForm({ ...prop }); setEditando(false); }} style={{ padding: "6px 14px", background: "#2d2b55", border: "1px solid #3d3b6e", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Cancelar</button>
              </>
            ) : confirmandoBorrado ? (
              <>
                <span style={{ fontSize: 12, color: "#f38ba8", alignSelf: "center", marginRight: 4 }}>¿Eliminar definitivamente?</span>
                <button onClick={() => onDelete(prop.id)} style={{ padding: "6px 14px", background: "#f38ba8", color: "#e0e0f0", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Sí, eliminar</button>
                <button onClick={() => setConfirmandoBorrado(false)} style={{ padding: "6px 14px", background: "#2d2b55", border: "1px solid #3d3b6e", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Cancelar</button>
              </>
            ) : (
              <>
                <button onClick={() => setEditando(true)} style={{ padding: "6px 14px", background: "#2d2b55", border: "1px solid #3d3b6e", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Editar</button>
                <button onClick={() => setConfirmandoBorrado(true)} style={{ padding: "6px 14px", background: "#3a1e1e", color: "#f38ba8", border: "1px solid #f0c9c9", borderRadius: 6, cursor: "pointer", fontSize: 13 }}>Eliminar</button>
              </>
            )}
            <button onClick={onClose} style={{ padding: "6px 10px", background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#c0c0e0" }}>×</button>
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 24px" }}>
          <div>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#c0c0e0", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10, marginTop: 4 }}>Contrato</div>
            {campo("Inquilino", "inquilino")}
            {campo("DNI / CUIT", "dni")}
            {campo("Período", "periodo")}
            {campo("Inmobiliaria", "inmobiliaria")}
            {campo("Comisión", "comision")}
            {campo("Teléfono", "telefono")}
            {campo("Mail", "mail")}
            {campo("Depósito", "deposito")}
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 600, color: "#c0c0e0", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10, marginTop: 4 }}>Propiedad</div>
            {campo("M²", "mts2")}
            {campo("Ambientes", "ambientes")}
            {campo("Cochera", "cochera")}
            {campo("Baulera", "baulera")}
            {campo("Alquiler 1° año", "alq_1")}
            {campo("Alquiler 2° año", "alq_2")}
            {campo("Expensas", "expensas")}
            {campo("Valor compra U$S", "valor_compra")}
            {campo("Valor mercado", "valor_mercado")}
            <div style={{ fontSize: 12, fontWeight: 600, color: "#c0c0e0", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10, marginTop: 16 }}>Administración</div>
            {campo("Fecha de pago", "fecha_pago")}
            {campo("Impuesto inmobiliario", "impuesto")}
            {campo("N° partida / ABL", "nro_partida")}
            {campo("Adm. edificio", "admin_edificio")}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL NUEVO COBRO (con carga por foto) ──────────────────────────────────
function ModalCobro({ prop, onClose, onGuardar }) {
  const hoy = new Date().toISOString().split("T")[0];
  const [form, setForm] = useState({ fecha: hoy, monto: prop?.alq_1 || "", moneda: "ARS", mes: "", anio: new Date().getFullYear(), concepto: "Alquiler", notas: "" });
  const [analizando, setAnalizando] = useState(false);
  const [errorFoto, setErrorFoto] = useState("");
  const [fotoNombre, setFotoNombre] = useState("");
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async function handleFoto(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setFotoNombre(file.name);
    setErrorFoto("");
    setAnalizando(true);
    try {
      const base64 = await fileToBase64(file);
      const mediaType = file.type || "image/jpeg";

      const response = await fetch("/api/leer-comprobante", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ base64, mediaType }),
      });

      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Error al leer el comprobante");

      const textBlock = data.content?.find((b) => b.type === "text");
      if (!textBlock) throw new Error("Sin respuesta de texto");

      let cleaned = textBlock.text.trim().replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(cleaned);

      setForm((f) => ({
        ...f,
        monto: parsed.monto != null ? String(parsed.monto) : f.monto,
        moneda: parsed.moneda || f.moneda,
        fecha: parsed.fecha || f.fecha,
        notas: parsed.notas || f.notas,
      }));
    } catch (err) {
      console.error(err);
      setErrorFoto("No se pudo leer el comprobante automáticamente. Completá los datos a mano.");
    } finally {
      setAnalizando(false);
    }
  }

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#252538", borderRadius: 16, width: "min(480px, 96vw)", padding: 28, maxHeight: "92vh", overflow: "auto" }}>
        <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 4 }}>Registrar cobro</div>
        <div style={{ fontSize: 13, color: "#c0c0e0", marginBottom: 16 }}>{prop?.direccion}</div>

        <label style={{ display: "block", border: "1.5px dashed #c8e0d0", borderRadius: 10, padding: "14px 12px", textAlign: "center", cursor: "pointer", background: "#252538", marginBottom: 16 }}>
          <input type="file" accept="image/*" onChange={handleFoto} style={{ display: "none" }} />
          {analizando ? (
            <span style={{ fontSize: 13, color: "#7c6fcd", fontWeight: 500 }}>Leyendo comprobante...</span>
          ) : fotoNombre ? (
            <span style={{ fontSize: 13, color: "#7c6fcd" }}>✓ {fotoNombre} — datos completados, revisá antes de guardar</span>
          ) : (
            <span style={{ fontSize: 13, color: "#7c6fcd", fontWeight: 500 }}>📷 Subir foto del comprobante (opcional)</span>
          )}
        </label>
        {errorFoto && <div style={{ fontSize: 12, color: "#f38ba8", marginBottom: 12, marginTop: -8 }}>{errorFoto}</div>}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Concepto</label>
            <select value={form.concepto} onChange={(e) => set("concepto", e.target.value)} style={{ width: "100%", padding: "7px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }}>
              <option>Alquiler</option><option>Expensas</option><option>Depósito</option><option>Otro</option>
            </select>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Fecha de cobro</label>
            <input type="date" value={form.fecha} onChange={(e) => set("fecha", e.target.value)} style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }} />
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Mes correspondiente</label>
            <select value={form.mes} onChange={(e) => set("mes", e.target.value)} style={{ width: "100%", padding: "7px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }}>
              <option value="">— Mes —</option>
              {meses.map((m, i) => <option key={i} value={m}>{m}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Año</label>
            <input type="number" value={form.anio} onChange={(e) => set("anio", e.target.value)} style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }} />
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Monto</label>
            <input type="number" value={form.monto} onChange={(e) => set("monto", e.target.value)} style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }} />
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Moneda</label>
            <select value={form.moneda} onChange={(e) => set("moneda", e.target.value)} style={{ width: "100%", padding: "7px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }}>
              <option>ARS</option><option>USD</option>
            </select>
          </div>
        </div>
        <div style={{ marginTop: 12 }}>
          <label style={{ fontSize: 12, color: "#c0c0e0" }}>Notas</label>
          <input type="text" value={form.notas} onChange={(e) => set("notas", e.target.value)} placeholder="Transferencia, efectivo, op. N°..." style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 4 }} />
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 20, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={{ padding: "8px 16px", background: "#2d2b55", border: "1px solid #3d3b6e", borderRadius: 6, cursor: "pointer" }}>Cancelar</button>
          <button onClick={() => onGuardar({ ...form, prop_id: prop.id, direccion: prop.direccion, inquilino: prop.inquilino })} style={{ padding: "8px 16px", background: "#7c6fcd", color: "#e0e0f0", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
            Registrar cobro
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── MODAL NUEVA PROPIEDAD ────────────────────────────────────────────────────
function ModalNueva({ onClose, onGuardar }) {
  const [form, setForm] = useState({ direccion: "", titular: "GC", estado: "libre", inquilino: "", dni: "", periodo: "", inmobiliaria: "", comision: "0.05", telefono: "", mail: "", alq_1: "", alq_2: "", expensas: "", mts2: "", ambientes: "", cochera: "", deposito: "", nro_partida: "", admin_edificio: "", valor_compra: "" });
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const inp = (label, key, tipo = "text", ph = "") => (
    <div>
      <label style={{ fontSize: 12, color: "#c0c0e0" }}>{label}</label>
      <input type={tipo} value={form[key] || ""} onChange={(e) => set(key, e.target.value)} placeholder={ph} style={{ width: "100%", padding: "6px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 3 }} />
    </div>
  );
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: "#252538", borderRadius: 16, width: "min(680px, 96vw)", maxHeight: "90vh", overflow: "auto", padding: 28 }}>
        <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 20 }}>Nueva propiedad</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <div style={{ gridColumn: "1 / -1" }}>{inp("Dirección *", "direccion", "text", "Ej: Gurruchaga 381 - 5A - CABA")}</div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Titular</label>
            <select value={form.titular} onChange={(e) => set("titular", e.target.value)} style={{ width: "100%", padding: "7px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 3 }}>
              <option value="GC">GC – Gerardo Castillo</option>
              <option value="EC">EC – Eliana Castillo</option>
              <option value="CC">CC – Cecilia Castillo</option>
              <option value="EC-CC-GC">EC-CC-GC (conjunto)</option>
            </select>
          </div>
          <div>
            <label style={{ fontSize: 12, color: "#c0c0e0" }}>Estado</label>
            <select value={form.estado} onChange={(e) => set("estado", e.target.value)} style={{ width: "100%", padding: "7px 8px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, marginTop: 3 }}>
              <option value="alquilado">Alquilado</option>
              <option value="libre">Libre</option>
              <option value="vendido">Vendido</option>
            </select>
          </div>
          {inp("Inquilino", "inquilino")}
          {inp("DNI / CUIT", "dni")}
          {inp("Período (DD/MM/AAAA A DD/MM/AAAA)", "periodo", "text", "01/01/2025 A 31/12/2026")}
          {inp("Inmobiliaria", "inmobiliaria")}
          {inp("Alquiler 1° año (ARS)", "alq_1", "number")}
          {inp("Alquiler 2° año (ARS)", "alq_2", "number")}
          {inp("Expensas (ARS)", "expensas", "number")}
          {inp("M²", "mts2", "number")}
          {inp("Ambientes", "ambientes")}
          {inp("Cochera", "cochera")}
          {inp("Depósito", "deposito")}
          {inp("Valor de compra U$S", "valor_compra", "number")}
          {inp("N° Partida / ABL", "nro_partida")}
          <div style={{ gridColumn: "1 / -1" }}>{inp("Adm. edificio", "admin_edificio")}</div>
          {inp("Teléfono inquilino", "telefono")}
          {inp("Mail inquilino", "mail")}
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 24, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={{ padding: "8px 16px", background: "#2d2b55", border: "1px solid #3d3b6e", borderRadius: 6, cursor: "pointer" }}>Cancelar</button>
          <button onClick={() => { if (!form.direccion) return alert("La dirección es obligatoria"); onGuardar(form); }} style={{ padding: "8px 16px", background: "#7c6fcd", color: "#e0e0f0", border: "none", borderRadius: 6, cursor: "pointer", fontWeight: 600 }}>
            Guardar propiedad
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── PESTAÑA COBROS con filtros, totales e impresión ─────────────────────────
function CobroTab({ cobros }) {
  const [filtroInq, setFiltroInq] = useState("");
  const [filtroDesde, setFiltroDesde] = useState("");
  const [filtroHasta, setFiltroHasta] = useState("");
  const [agrupar, setAgrupar] = useState("ninguno");

  const inquilinos = [...new Set(cobros.map(c => c.inquilino).filter(Boolean))].sort();

  const filtrados = cobros.filter(c => {
    const okInq = !filtroInq || c.inquilino === filtroInq;
    const okDesde = !filtroDesde || c.fecha >= filtroDesde;
    const okHasta = !filtroHasta || c.fecha <= filtroHasta;
    return okInq && okDesde && okHasta;
  });

  const totalARS = filtrados.filter(c => c.moneda === "ARS").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);
  const totalUSD = filtrados.filter(c => c.moneda === "USD").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);

  // Agrupación
  let grupos = null;
  if (agrupar === "inquilino") {
    const map = {};
    filtrados.forEach(c => {
      const k = c.inquilino || "Sin nombre";
      if (!map[k]) map[k] = [];
      map[k].push(c);
    });
    grupos = Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  } else if (agrupar === "mes") {
    const map = {};
    filtrados.forEach(c => {
      const k = `${c.anio}-${c.mes || "Sin mes"}`;
      if (!map[k]) map[k] = [];
      map[k].push(c);
    });
    grupos = Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
  }

  function imprimirPDF() {
    const style = `
      <style>
        body { font-family: Arial, sans-serif; font-size: 12px; color: #1a1a2e; margin: 20px; }
        h2 { color: #5a4fa3; margin-bottom: 4px; }
        .sub { color: #888; margin-bottom: 16px; font-size: 11px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th { background: #5a4fa3; color: white; padding: 7px 10px; text-align: left; font-size: 11px; text-transform: uppercase; }
        td { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; }
        tr:nth-child(even) { background: #f5f4ff; }
        .grupo-titulo { font-weight: bold; color: #5a4fa3; margin: 16px 0 4px; font-size: 13px; border-bottom: 2px solid #5a4fa3; padding-bottom: 4px; }
        .total-row { background: #ece9ff !important; font-weight: bold; }
        .totales { background: #5a4fa3; color: white; padding: 10px 14px; border-radius: 6px; margin-bottom: 16px; display: flex; gap: 24px; }
        .totales span { font-size: 13px; }
        @media print { body { margin: 10px; } }
      </style>`;

    const fecha = new Date().toLocaleDateString("es-AR");
    let html = `${style}<h2>Reporte de Cobros</h2>
      <div class="sub">Generado el ${fecha}${filtroInq ? ` · Inquilino: ${filtroInq}` : ""}${filtroDesde ? ` · Desde: ${filtroDesde}` : ""}${filtroHasta ? ` · Hasta: ${filtroHasta}` : ""}</div>
      <div class="totales">
        <span>Total ARS: <b>$${totalARS.toLocaleString("es-AR")}</b></span>
        ${totalUSD > 0 ? `<span>Total USD: <b>U$S ${totalUSD.toLocaleString()}</b></span>` : ""}
        <span>Registros: <b>${filtrados.length}</b></span>
      </div>`;

    const filasCobros = (lista) => lista.map(c => `
      <tr>
        <td>${c.fecha ? new Date(c.fecha).toLocaleDateString("es-AR") : "—"}</td>
        <td>${c.inquilino || "—"}</td>
        <td>${c.direccion || "—"}</td>
        <td>${c.concepto || "—"}</td>
        <td>${c.mes || ""} ${c.anio || ""}</td>
        <td><b>${c.moneda === "ARS" ? "$" + Number(c.monto).toLocaleString("es-AR") : "U$S " + Number(c.monto).toLocaleString()}</b></td>
        <td>${c.notas || "—"}</td>
      </tr>`).join("");

    const encabezado = `<tr><th>Fecha</th><th>Inquilino</th><th>Dirección</th><th>Concepto</th><th>Mes/Año</th><th>Monto</th><th>Notas</th></tr>`;

    if (grupos) {
      grupos.forEach(([grupo, lista]) => {
        const subARS = lista.filter(c => c.moneda === "ARS").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);
        const subUSD = lista.filter(c => c.moneda === "USD").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);
        html += `<div class="grupo-titulo">${grupo}</div>
          <table>${encabezado}${filasCobros(lista)}
            <tr class="total-row">
              <td colspan="5">Subtotal ${grupo}</td>
              <td>$${subARS.toLocaleString("es-AR")}${subUSD > 0 ? ` / U$S ${subUSD.toLocaleString()}` : ""}</td>
              <td>${lista.length} cobros</td>
            </tr>
          </table>`;
      });
    } else {
      html += `<table>${encabezado}${filasCobros(filtrados)}</table>`;
    }

    const win = window.open("", "_blank");
    win.document.write(`<html><head><title>Reporte de Cobros</title></head><body>${html}</body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 500);
  }

  const S = {
    inp: { padding: "7px 10px", border: "1px solid #5a4fa3", borderRadius: 6, fontSize: 13, background: "#1e1e2e", color: "#e0e0f0", outline: "none" },
    th: { padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase", background: "#2d2b55", borderBottom: "2px solid #5a4fa3" },
    td: { padding: "9px 14px", borderBottom: "1px solid #3d3b6e", color: "#e0e0f0" },
    tdAlt: { padding: "9px 14px", borderBottom: "1px solid #3d3b6e", color: "#e0e0f0", background: "#252545" },
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ fontSize: 20, fontWeight: 600, flex: 1, color: "#e0e0f0" }}>Registro de cobros</div>
        <button onClick={imprimirPDF} style={{ padding: "7px 16px", background: "#7c6fcd", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
          Imprimir / PDF
        </button>
      </div>

      {/* FILTROS */}
      <div style={{ background: "#252538", border: "1px solid #5a4fa3", borderRadius: 10, padding: "14px 16px", marginBottom: 16, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
        <div>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 4, fontWeight: 600 }}>INQUILINO</div>
          <select value={filtroInq} onChange={e => setFiltroInq(e.target.value)} style={S.inp}>
            <option value="">Todos</option>
            {inquilinos.map(i => <option key={i} value={i}>{i}</option>)}
          </select>
        </div>
        <div>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 4, fontWeight: 600 }}>DESDE</div>
          <input type="date" value={filtroDesde} onChange={e => setFiltroDesde(e.target.value)} style={S.inp} />
        </div>
        <div>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 4, fontWeight: 600 }}>HASTA</div>
          <input type="date" value={filtroHasta} onChange={e => setFiltroHasta(e.target.value)} style={S.inp} />
        </div>
        <div>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 4, fontWeight: 600 }}>AGRUPAR POR</div>
          <select value={agrupar} onChange={e => setAgrupar(e.target.value)} style={S.inp}>
            <option value="ninguno">Sin agrupar</option>
            <option value="inquilino">Inquilino</option>
            <option value="mes">Mes</option>
          </select>
        </div>
        <button onClick={() => { setFiltroInq(""); setFiltroDesde(""); setFiltroHasta(""); setAgrupar("ninguno"); }}
          style={{ padding: "7px 12px", background: "none", border: "1px solid #5a4fa3", borderRadius: 6, color: "#cba6f7", cursor: "pointer", fontSize: 12 }}>
          Limpiar
        </button>
      </div>

      {/* TOTALES */}
      <div style={{ display: "flex", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ background: "#2d2b55", border: "1px solid #5a4fa3", borderRadius: 8, padding: "10px 18px" }}>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 2 }}>TOTAL ARS</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#a6e3a1" }}>${totalARS.toLocaleString("es-AR")}</div>
        </div>
        {totalUSD > 0 && (
          <div style={{ background: "#2d2b55", border: "1px solid #5a4fa3", borderRadius: 8, padding: "10px 18px" }}>
            <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 2 }}>TOTAL USD</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: "#89b4fa" }}>U$S {totalUSD.toLocaleString()}</div>
          </div>
        )}
        <div style={{ background: "#2d2b55", border: "1px solid #5a4fa3", borderRadius: 8, padding: "10px 18px" }}>
          <div style={{ fontSize: 11, color: "#cba6f7", marginBottom: 2 }}>REGISTROS</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#e0e0f0" }}>{filtrados.length}</div>
        </div>
      </div>

      {cobros.length === 0 ? (
        <div style={{ background: "#252538", borderRadius: 10, border: "1px solid #3d3b6e", padding: 48, textAlign: "center", color: "#c0c0e0" }}>
          <div style={{ fontSize: 15, marginBottom: 8, color: "#d0d0f0" }}>Sin cobros registrados</div>
          <div style={{ fontSize: 13 }}>Desde Propiedades, hacé clic en "+ Cobro" en cualquier unidad alquilada.</div>
        </div>
      ) : grupos ? (
        grupos.map(([grupo, lista]) => {
          const subARS = lista.filter(c => c.moneda === "ARS").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);
          const subUSD = lista.filter(c => c.moneda === "USD").reduce((s, c) => s + (parseFloat(c.monto) || 0), 0);
          return (
            <div key={grupo} style={{ marginBottom: 24 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: "#cba6f7", marginBottom: 8, borderBottom: "2px solid #5a4fa3", paddingBottom: 6 }}>
                {grupo}
                <span style={{ fontSize: 12, fontWeight: 400, color: "#c0c0e0", marginLeft: 12 }}>
                  {lista.length} cobros · ${subARS.toLocaleString("es-AR")} ARS{subUSD > 0 ? ` · U$S ${subUSD.toLocaleString()}` : ""}
                </span>
              </div>
              <div style={{ background: "#252538", borderRadius: 8, border: "1px solid #3d3b6e", overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead><tr>{["Fecha","Inquilino","Dirección","Concepto","Mes/Año","Monto","Notas"].map(h => <th key={h} style={S.th}>{h}</th>)}</tr></thead>
                  <tbody>
                    {lista.map((c, i) => (
                      <tr key={c.id}>
                        <td style={i%2===0?S.td:S.tdAlt}>{c.fecha ? new Date(c.fecha).toLocaleDateString("es-AR") : "—"}</td>
                        <td style={i%2===0?S.td:S.tdAlt}>{c.inquilino || "—"}</td>
                        <td style={{...(i%2===0?S.td:S.tdAlt), maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{c.direccion}</td>
                        <td style={i%2===0?S.td:S.tdAlt}><span style={{background:"#3d3b6e",color:"#cba6f7",borderRadius:4,padding:"1px 7px",fontSize:11}}>{c.concepto}</span></td>
                        <td style={i%2===0?S.td:S.tdAlt}>{c.mes} {c.anio}</td>
                        <td style={{...(i%2===0?S.td:S.tdAlt), fontWeight:600, color:"#a6e3a1"}}>{c.moneda==="ARS"?fmtARS(c.monto):`U$S ${Number(c.monto).toLocaleString()}`}</td>
                        <td style={{...(i%2===0?S.td:S.tdAlt), color:"#a0a0c0"}}>{c.notas||"—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          );
        })
      ) : (
        <div style={{ background: "#252538", borderRadius: 10, border: "1px solid #3d3b6e", overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr>{["Fecha","Inquilino","Dirección","Concepto","Mes/Año","Monto","Notas"].map(h => <th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {filtrados.map((c, i) => (
                <tr key={c.id}>
                  <td style={i%2===0?S.td:S.tdAlt}>{c.fecha ? new Date(c.fecha).toLocaleDateString("es-AR") : "—"}</td>
                  <td style={i%2===0?S.td:S.tdAlt}>{c.inquilino || "—"}</td>
                  <td style={{...(i%2===0?S.td:S.tdAlt), maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{c.direccion}</td>
                  <td style={i%2===0?S.td:S.tdAlt}><span style={{background:"#3d3b6e",color:"#cba6f7",borderRadius:4,padding:"1px 7px",fontSize:11}}>{c.concepto}</span></td>
                  <td style={i%2===0?S.td:S.tdAlt}>{c.mes} {c.anio}</td>
                  <td style={{...(i%2===0?S.td:S.tdAlt), fontWeight:600, color:"#a6e3a1"}}>{c.moneda==="ARS"?fmtARS(c.monto):`U$S ${Number(c.monto).toLocaleString()}`}</td>
                  <td style={{...(i%2===0?S.td:S.tdAlt), color:"#a0a0c0"}}>{c.notas||"—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ─── PESTAÑA PLANILLA: vista tipo Excel con todas las columnas ───────────────
const COLUMNAS_PLANILLA = [
  { key: "unidad", label: "Unidad", w: 60 },
  { key: "titular", label: "Titular", w: 70 },
  { key: "direccion", label: "Dirección", w: 220 },
  { key: "estado", label: "Estado", w: 90 },
  { key: "inquilino", label: "Inquilino", w: 180 },
  { key: "dni", label: "DNI/CUIT", w: 110 },
  { key: "periodo", label: "Período", w: 170 },
  { key: "inmobiliaria", label: "Inmobiliaria", w: 130 },
  { key: "comision", label: "Comisión", w: 80 },
  { key: "alq_1", label: "Alq. 1° año", w: 100 },
  { key: "alq_2", label: "Alq. 2° año", w: 100 },
  { key: "expensas", label: "Expensas", w: 90 },
  { key: "fecha_pago", label: "Fecha pago", w: 130 },
  { key: "mts2", label: "M²", w: 60 },
  { key: "ambientes", label: "Ambientes", w: 80 },
  { key: "cochera", label: "Cochera", w: 90 },
  { key: "baulera", label: "Baulera", w: 90 },
  { key: "valor_compra", label: "Val. compra", w: 100 },
  { key: "valor_mercado", label: "Val. mercado", w: 110 },
  { key: "deposito", label: "Depósito", w: 110 },
  { key: "telefono", label: "Teléfono", w: 120 },
  { key: "mail", label: "Mail", w: 180 },
  { key: "impuesto", label: "Impuesto inmob.", w: 160 },
  { key: "nro_partida", label: "N° Partida", w: 140 },
  { key: "admin_edificio", label: "Adm. edificio", w: 200 },
];

function PlanillaTab({ props, onUpdateCell }) {
  const [busqueda, setBusqueda] = useState("");
  const [celda, setCelda] = useState({ fila: 0, col: 0 }); // celda activa
  const [editando, setEditando] = useState(false);
  const [valEdit, setValEdit] = useState("");
  const tableRef = useRef(null);

  const filtradas = props.filter((p) => {
    const txt = (p.direccion + " " + p.inquilino).toLowerCase();
    return !busqueda || txt.includes(busqueda.toLowerCase());
  });

  const COLS = COLUMNAS_PLANILLA;
  const nCols = COLS.length;
  const nFilas = filtradas.length;

  function mover(df, dc) {
    setCelda(prev => ({
      fila: Math.max(0, Math.min(nFilas - 1, prev.fila + df)),
      col: Math.max(0, Math.min(nCols - 1, prev.col + dc)),
    }));
  }

  function iniciarEdicion(fila, col) {
    const p = filtradas[fila];
    const key = COLS[col].key;
    if (key === "estado") return; // estado se edita con select directo
    setValEdit(p[key] || "");
    setEditando(true);
  }

  function confirmarEdicion() {
    if (!editando) return;
    const p = filtradas[celda.fila];
    const key = COLS[celda.col].key;
    const valorActual = p[key] || "";
    if (valEdit !== valorActual) onUpdateCell(p.id, key, valEdit);
    setEditando(false);
  }

  function handleKeyDown(e) {
    if (editando) {
      if (e.key === "Escape") { setEditando(false); e.preventDefault(); }
      if (e.key === "Enter") { confirmarEdicion(); e.preventDefault(); }
      if (e.key === "Tab") { confirmarEdicion(); mover(0, e.shiftKey ? -1 : 1); e.preventDefault(); }
      return;
    }
    switch (e.key) {
      case "ArrowUp":    mover(-1, 0); e.preventDefault(); break;
      case "ArrowDown":  mover(1, 0);  e.preventDefault(); break;
      case "ArrowLeft":  mover(0, -1); e.preventDefault(); break;
      case "ArrowRight": mover(0, 1);  e.preventDefault(); break;
      case "Tab":        mover(0, e.shiftKey ? -1 : 1); e.preventDefault(); break;
      case "Enter":      iniciarEdicion(celda.fila, celda.col); e.preventDefault(); break;
      case "F2":         iniciarEdicion(celda.fila, celda.col); e.preventDefault(); break;
      default:
        // Empezar a escribir directamente activa la edición
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey) {
          setValEdit(e.key);
          setEditando(true);
          e.preventDefault();
        }
    }
  }

  // Scroll automático a la celda activa
  useEffect(() => {
    if (!tableRef.current) return;
    const td = tableRef.current.querySelector(`[data-celda="${celda.fila}-${celda.col}"]`);
    if (td) td.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [celda]);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
        <div style={{ fontSize: 20, fontWeight: 600, flex: 1 }}>Planilla completa</div>
        <input value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar..." style={{ padding: "7px 12px", border: "1px solid #3d3b6e", borderRadius: 6, fontSize: 13, width: 220, background: "#1e1e2e", color: "#e0e0f0" }} />
      </div>
      <div style={{ fontSize: 12, color: "#c0c0e0", marginBottom: 10 }}>
        Usá las <b style={{color:"#cba6f7"}}>flechas del teclado</b> para moverte · <b style={{color:"#cba6f7"}}>Enter o F2</b> para editar · <b style={{color:"#cba6f7"}}>Escape</b> para cancelar · También podés hacer clic en cualquier celda
      </div>
      <div
        ref={tableRef}
        tabIndex={0}
        onKeyDown={handleKeyDown}
        onFocus={() => {}}
        style={{ background: "#fff", borderRadius: 10, border: "1px solid #3d3b6e", overflow: "auto", maxWidth: "100%", maxHeight: "70vh", outline: "none" }}
      >
        <table style={{ borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
          <thead>
            <tr style={{ background: "#2d2b55" }}>
              {COLS.map((c) => (
                <th key={c.key} style={{ position: "sticky", top: 0, zIndex: 2, background: "#2d2b55", padding: "8px 6px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 10.5, textTransform: "uppercase", borderBottom: "2px solid #5a4fa3", borderRight: "1px solid #3d3b6e", width: c.w, minWidth: c.w }}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtradas.map((p, fi) => (
              <tr key={p.id} style={{ background: fi % 2 === 0 ? "#fff" : "#f5f4ff" }}>
                {COLS.map((c, ci) => {
                  const activa = celda.fila === fi && celda.col === ci;
                  const estaEditando = activa && editando;
                  return (
                    <td
                      key={c.key}
                      data-celda={`${fi}-${ci}`}
                      onClick={() => { setCelda({ fila: fi, col: ci }); tableRef.current?.focus(); }}
                      onDoubleClick={() => { setCelda({ fila: fi, col: ci }); iniciarEdicion(fi, ci); }}
                      style={{
                        borderRight: "1px solid #e0e0e0",
                        borderBottom: "1px solid #e8e8f0",
                        padding: 0,
                        outline: activa ? "2px solid #7c6fcd" : "none",
                        outlineOffset: "-2px",
                        background: activa ? (estaEditando ? "#000" : "#ede9ff") : "transparent",
                        position: "relative",
                      }}
                    >
                      {c.key === "estado" ? (
                        <select
                          value={p.estado}
                          onChange={(e) => onUpdateCell(p.id, "estado", e.target.value)}
                          style={{ width: "100%", border: "none", background: "transparent", fontSize: 12, padding: "5px 4px", fontFamily: "inherit", color: "#111", cursor: "pointer" }}
                        >
                          <option value="alquilado">Alquilado</option>
                          <option value="libre">Libre</option>
                          <option value="vendido">Vendido</option>
                        </select>
                      ) : estaEditando ? (
                        <input
                          autoFocus
                          value={valEdit}
                          onChange={(e) => setValEdit(e.target.value)}
                          onBlur={confirmarEdicion}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") { confirmarEdicion(); e.preventDefault(); }
                            if (e.key === "Escape") { setEditando(false); e.preventDefault(); tableRef.current?.focus(); }
                            if (e.key === "Tab") { confirmarEdicion(); mover(0, e.shiftKey ? -1 : 1); e.preventDefault(); tableRef.current?.focus(); }
                            e.stopPropagation();
                          }}
                          style={{ width: "100%", border: "none", padding: "4px 6px", fontSize: 12, fontFamily: "inherit", background: "#000", color: "#fff", outline: "none" }}
                        />
                      ) : (
                        <div style={{ padding: "5px 6px", minHeight: 18, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", color: "#111" }}>
                          {p[c.key] || <span style={{ color: "#ccc" }}>—</span>}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 8, fontSize: 12, color: "#c0c0e0" }}>{filtradas.length} de {props.length} propiedades · {COLS.length} columnas · Celda: fila {celda.fila + 1}, col {celda.col + 1}</div>
    </div>
  );
}


// ─── APP PRINCIPAL ────────────────────────────────────────────────────────────
export default function App() {
  const [props, setProps] = useState([]);
  const [cobros, setCobros] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [busqueda, setBusqueda] = useState("");
  const [filtroEstado, setFiltroEstado] = useState("todos");
  const [filtroTitular, setFiltroTitular] = useState("todos");
  const [propSel, setPropSel] = useState(null);
  const [cobroModal, setCobroModal] = useState(null);
  const [nuevaModal, setNuevaModal] = useState(false);
  const [filtroCobros, setFiltroCobros] = useState("");

  // ── Carga inicial: trae datos de Supabase, y si está vacío, migra el Excel ──
  useEffect(() => {
    async function cargar() {
      setCargando(true);
      const { data: propsData, error: e1 } = await supabase.from("propiedades").select("*").order("direccion");
      const { data: cobrosData, error: e2 } = await supabase.from("cobros").select("*").order("fecha", { ascending: false });

      if (e1 || e2) {
        console.error(e1 || e2);
        setCargando(false);
        return;
      }

      if (propsData.length === 0) {
        // Primera vez: migrar datos del Excel a Supabase
        const { data: inserted, error: insErr } = await supabase.from("propiedades").insert(PROPIEDADES_INICIALES).select();
        if (!insErr) setProps(inserted);
      } else {
        setProps(propsData);
      }
      setCobros(cobrosData || []);
      setCargando(false);
    }
    cargar();
  }, []);

  const propsFiltradas = props.filter((p) => {
    const txt = (p.direccion + " " + p.inquilino + " " + p.titular).toLowerCase();
    const ok_bus = !busqueda || txt.includes(busqueda.toLowerCase());
    const sem = semaforo(p);
    const ok_est = filtroEstado === "todos" || sem === filtroEstado;
    const ok_tit = filtroTitular === "todos" || p.titular.startsWith(filtroTitular);
    return ok_bus && ok_est && ok_tit;
  });

  const stats = {
    total: props.length,
    alquiladas: props.filter((p) => p.estado === "alquilado").length,
    libres: props.filter((p) => p.estado === "libre").length,
    vendidas: props.filter((p) => p.estado === "vendido").length,
    vencidos: props.filter((p) => semaforo(p) === "vencido").length,
    proximos: props.filter((p) => semaforo(p) === "proximo").length,
    cobradoMes: cobros.filter((c) => {
      const d = new Date(c.fecha);
      const hoy = new Date();
      return d.getMonth() === hoy.getMonth() && d.getFullYear() === hoy.getFullYear() && c.moneda === "ARS";
    }).reduce((s, c) => s + (parseFloat(c.monto) || 0), 0),
  };

  async function guardarProp(updated) {
    const { id, ...rest } = updated;
    const { error } = await supabase.from("propiedades").update(rest).eq("id", id);
    if (!error) {
      setProps((ps) => ps.map((p) => (p.id === id ? updated : p)));
    } else {
      alert("Error al guardar: " + error.message);
    }
    setPropSel(null);
  }

  async function actualizarCelda(id, campo, valor) {
    const { error } = await supabase.from("propiedades").update({ [campo]: valor }).eq("id", id);
    if (!error) {
      setProps((ps) => ps.map((p) => (p.id === id ? { ...p, [campo]: valor } : p)));
    } else {
      alert("Error al guardar: " + error.message);
    }
  }

  async function eliminarProp(id) {
    const { error } = await supabase.from("propiedades").delete().eq("id", id);
    if (!error) {
      setProps((ps) => ps.filter((p) => p.id !== id));
      setPropSel(null);
    } else {
      alert("Error al eliminar: " + error.message);
    }
  }

  async function guardarCobro(cobro) {
    const { data, error } = await supabase.from("cobros").insert([cobro]).select();
    if (!error) {
      setCobros((cs) => [data[0], ...cs]);
    } else {
      alert("Error al registrar cobro: " + error.message);
    }
    setCobroModal(null);
  }

  async function guardarNueva(nueva) {
    const { data, error } = await supabase.from("propiedades").insert([nueva]).select();
    if (!error) {
      setProps((ps) => [...ps, data[0]]);
    } else {
      alert("Error al crear propiedad: " + error.message);
    }
    setNuevaModal(false);
  }

  function exportarExcel() {
    const wb = XLSX.utils.book_new();

    const propsExport = props.map((p) => ({
      ID: p.id, Unidad: p.unidad, Titular: p.titular, Dirección: p.direccion, Estado: p.estado,
      Inquilino: p.inquilino, DNI: p.dni, Período: p.periodo, Inmobiliaria: p.inmobiliaria,
      Comisión: p.comision, "Alquiler 1° año": p.alq_1, "Alquiler 2° año": p.alq_2, Expensas: p.expensas,
      "Fecha de pago": p.fecha_pago, "M2": p.mts2, Ambientes: p.ambientes, Cochera: p.cochera, Baulera: p.baulera,
      "Valor compra": p.valor_compra, "Valor mercado": p.valor_mercado, Depósito: p.deposito,
      Teléfono: p.telefono, Mail: p.mail, "Impuesto inmobiliario": p.impuesto, "N° Partida": p.nro_partida,
      "Adm. edificio": p.admin_edificio,
    }));
    const wsProps = XLSX.utils.json_to_sheet(propsExport);
    XLSX.utils.book_append_sheet(wb, wsProps, "Propiedades");

    const cobrosExport = cobros.map((c) => ({
      ID: c.id, "ID Propiedad": c.prop_id, Fecha: c.fecha, Dirección: c.direccion, Inquilino: c.inquilino,
      Concepto: c.concepto, Mes: c.mes, Año: c.anio, Monto: c.monto, Moneda: c.moneda, Notas: c.notas,
    }));
    const wsCobros = XLSX.utils.json_to_sheet(cobrosExport);
    XLSX.utils.book_append_sheet(wb, wsCobros, "Cobros");

    const wsInfo = XLSX.utils.aoa_to_sheet([
      ["COPIA DE SEGURIDAD — Gestión Inmobiliaria"],
      [""],
      ["Fecha de exportación:", new Date().toLocaleString("es-AR")],
      ["Total de propiedades:", props.length],
      ["Total de cobros registrados:", cobros.length],
      [""],
      ["Este archivo es un respaldo completo de los datos cargados en la app."],
      ["Si en algún momento hay un problema con los servidores (Vercel o Supabase),"],
      ["este Excel contiene toda la información necesaria para reconstruir el portfolio:"],
      ["- Hoja 'Propiedades': todos los datos de cada unidad"],
      ["- Hoja 'Cobros': historial completo de pagos registrados"],
      [""],
      ["Recomendación: descargar este backup una vez por semana o luego de cargar"],
      ["varios cobros nuevos, y guardarlo en Google Drive o Dropbox."],
    ]);
    XLSX.utils.book_append_sheet(wb, wsInfo, "Info del backup");

    const fecha = new Date().toISOString().split("T")[0];
    XLSX.writeFile(wb, `backup_inmuebles_${fecha}.xlsx`);
  }

  const NAV = [
    { id: "dashboard", label: "Dashboard" },
    { id: "propiedades", label: "Propiedades" },
    { id: "planilla", label: "Planilla" },
    { id: "cobros", label: "Cobros" },
    { id: "alertas", label: `Alertas${stats.vencidos + stats.proximos > 0 ? ` (${stats.vencidos + stats.proximos})` : ""}` },
  ];

  if (cargando) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "sans-serif", color: "#c0c0e0" }}>
        Cargando portfolio...
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "'Inter', -apple-system, sans-serif", minHeight: "100vh", background: "#1e1e2e", color: "#e0e0f0" }}>
      <div style={{ background: "#7c6fcd", padding: "0 24px", display: "flex", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
        <div style={{ padding: "14px 0", color: "#e0e0f0", fontWeight: 700, fontSize: 15, letterSpacing: "-0.01em", marginRight: 12 }}>
          Gestión Inmobiliaria
        </div>
        {NAV.map((n) => (
          <button key={n.id} onClick={() => setTab(n.id)}
            style={{ background: "none", border: "none", cursor: "pointer", padding: "16px 4px", fontSize: 13, color: tab === n.id ? "#fff" : "rgba(255,255,255,0.65)", borderBottom: tab === n.id ? "2px solid #fff" : "2px solid transparent", fontWeight: tab === n.id ? 600 : 400 }}>
            {n.label}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <button onClick={exportarExcel} style={{ padding: "7px 14px", background: "rgba(137,180,250,0.15)", color: "#e0e0f0", border: "1px solid rgba(137,180,250,0.3)", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 500 }}>
          Copia de seguridad
        </button>
        <button onClick={() => setNuevaModal(true)} style={{ padding: "7px 14px", background: "rgba(203,166,247,0.2)", color: "#e0e0f0", border: "1px solid rgba(203,166,247,0.4)", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 500 }}>
          + Nueva propiedad
        </button>
      </div>

      <div style={{ padding: "24px", maxWidth: 1200, margin: "0 auto" }}>
        {tab === "dashboard" && (
          <div>
            <div style={{ fontSize: 20, fontWeight: 600, marginBottom: 18 }}>Resumen del portfolio</div>
            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
              <KPI label="Total propiedades" value={stats.total} />
              <KPI label="Alquiladas" value={stats.alquiladas} color="#7c6fcd" />
              <KPI label="Libres" value={stats.libres} color="#89b4fa" />
              <KPI label="Vendidas" value={stats.vendidas} color="#888" />
              <KPI label="Contratos vencidos" value={stats.vencidos} color={stats.vencidos > 0 ? "#f38ba8" : "#e0e0f0"} sub="Requieren atención inmediata" />
              <KPI label="Vencen en 60 días" value={stats.proximos} color={stats.proximos > 0 ? "#fab387" : "#e0e0f0"} sub="Renovar pronto" />
              <KPI label="Cobrado este mes" value={"$" + stats.cobradoMes.toLocaleString("es-AR")} color="#7c6fcd" sub="Solo ARS" />
            </div>

            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 12 }}>Semáforo de vencimientos</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 10 }}>
              {props.filter((p) => p.estado !== "vendido" && p.estado !== "libre").sort((a, b) => {
                const da = diasParaVencer(a.periodo) ?? 9999;
                const db = diasParaVencer(b.periodo) ?? 9999;
                return da - db;
              }).slice(0, 18).map((p) => {
                const sem = semaforo(p);
                const dias = diasParaVencer(p.periodo);
                const c = COLORES_SEMAFORO[sem];
                return (
                  <div key={p.id} onClick={() => setPropSel(p)}
                    style={{ background: "#252538", border: `1px solid ${c.bg}`, borderLeft: `4px solid ${c.color}`, borderRadius: 8, padding: "12px 14px", cursor: "pointer" }}>
                    <div style={{ fontSize: 13, fontWeight: 500, marginBottom: 4 }}>{p.direccion}</div>
                    <div style={{ fontSize: 12, color: "#d0d0f0", marginBottom: 6 }}>{p.inquilino || "Sin inquilino"}</div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <Badge tipo={sem} />
                      {dias !== null && <span style={{ fontSize: 11, color: c.color, fontWeight: 500 }}>{dias < 0 ? `−${Math.abs(dias)}d` : `+${dias}d`}</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab === "propiedades" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 600, flex: 1 }}>Propiedades</div>
              <input value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar por dirección, inquilino..." style={{ padding: "7px 12px", border: "1px solid #5a4fa3", borderRadius: 6, fontSize: 13, width: 240, background: "#2d2b55", color: "#e0e0f0" }} />
              <select value={filtroEstado} onChange={(e) => setFiltroEstado(e.target.value)} style={{ padding: "7px 10px", border: "1px solid #5a4fa3", borderRadius: 6, fontSize: 13, background: "#2d2b55", color: "#e0e0f0" }}>
                <option value="todos">Todos los estados</option>
                <option value="vigente">Vigentes</option>
                <option value="proximo">Próximos a vencer</option>
                <option value="vencido">Vencidos</option>
                <option value="libre">Libres</option>
                <option value="vendido">Vendidos</option>
              </select>
              <select value={filtroTitular} onChange={(e) => setFiltroTitular(e.target.value)} style={{ padding: "7px 10px", border: "1px solid #5a4fa3", borderRadius: 6, fontSize: 13, background: "#2d2b55", color: "#e0e0f0" }}>
                <option value="todos">Todos los titulares</option>
                <option value="GC">GC – Gerardo</option>
                <option value="EC">EC – Eliana</option>
                <option value="CC">CC – Cecilia</option>
              </select>
            </div>
            <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #3d3b6e", overflow: "hidden" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: "#2d2b55", borderBottom: "2px solid #5a4fa3" }}>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Dirección</th>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Inquilino</th>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Titular</th>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Alquiler</th>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Vencimiento</th>
                    <th style={{ padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#cba6f7", fontSize: 11, textTransform: "uppercase" }}>Estado</th>
                    <th style={{ padding: "10px 14px" }}></th>
                  </tr>
                </thead>
                <tbody>
                  {propsFiltradas.map((p, i) => {
                    const sem = semaforo(p);
                    const fv = vencimientoContrato(p.periodo);
                    return (
                      <tr key={p.id} style={{ borderBottom: "1px solid #e8e8f0", background: i % 2 === 0 ? "#fff" : "#f5f4ff" }}
                        onMouseEnter={e => e.currentTarget.style.background = "#ede9ff"}
                        onMouseLeave={e => e.currentTarget.style.background = i % 2 === 0 ? "#fff" : "#f5f4ff"}>
                        <td style={{ padding: "10px 14px" }}>
                          <div style={{ fontWeight: 600, color: "#111" }}>{p.direccion}</div>
                          <div style={{ fontSize: 11, color: "#666", marginTop: 2 }}>{p.mts2 ? `${p.mts2} m² · ${p.ambientes} amb.` : ""}</div>
                        </td>
                        <td style={{ padding: "10px 14px", color: p.inquilino ? "#111" : "#aaa", fontWeight: p.inquilino ? 500 : 400 }}>{p.inquilino || "—"}</td>
                        <td style={{ padding: "10px 14px" }}><span style={{ background: "#2d2b55", color: "#cba6f7", borderRadius: 4, padding: "2px 8px", fontSize: 11, fontWeight: 700 }}>{p.titular}</span></td>
                        <td style={{ padding: "10px 14px", color: "#111", fontWeight: 600 }}>{p.alq_1 ? fmtARS(p.alq_1) : <span style={{color:"#aaa"}}>—</span>}</td>
                        <td style={{ padding: "10px 14px", color: "#111" }}>{fv ? fmtFecha(fv) : <span style={{ color: "#aaa" }}>—</span>}</td>
                        <td style={{ padding: "10px 14px" }}><Badge tipo={sem} /></td>
                        <td style={{ padding: "10px 14px" }}>
                          <div style={{ display: "flex", gap: 6 }}>
                            <button onClick={() => setPropSel(p)} style={{ padding: "4px 10px", border: "1px solid #7c6fcd", borderRadius: 5, background: "#fff", color: "#5a4fa3", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>Ver</button>
                            {p.estado === "alquilado" && (
                              <button onClick={() => setCobroModal(p)} style={{ padding: "4px 10px", border: "1px solid #5a4fa3", borderRadius: 5, background: "#2d2b55", color: "#cba6f7", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>+ Cobro</button>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {propsFiltradas.length === 0 && <div style={{ padding: 40, textAlign: "center", color: "#888" }}>Sin resultados</div>}
            </div>
            <div style={{ marginTop: 8, fontSize: 12, color: "#c0c0e0" }}>{propsFiltradas.length} de {props.length} propiedades</div>
          </div>
        )}

        {tab === "planilla" && (
          <PlanillaTab props={props} onUpdateCell={actualizarCelda} />
        )}

        {tab === "cobros" && (
          <CobroTab cobros={cobros} />
        )}

        {tab === "alertas" && (
          <div>
            <div style={{ fontSize: 20, fontWeight: 600, marginBottom: 18 }}>Alertas y atención requerida</div>
            {[
              { label: "Contratos vencidos", tipo: "vencido", desc: "El contrato ya venció. Contactar al inquilino para renovar o iniciar proceso de desalojo." },
              { label: "Vencen en los próximos 60 días", tipo: "proximo", desc: "Iniciar conversación de renovación con tiempo suficiente." },
              { label: "Propiedades libres", tipo: "libre", desc: "Sin inquilino activo. Evaluar poner en alquiler." },
              { label: "Sin fecha de contrato", tipo: "sin-fecha", desc: "Propiedades con contrato pero sin fecha de vencimiento registrada." },
            ].map((grupo) => {
              const lista = props.filter((p) => semaforo(p) === grupo.tipo);
              if (lista.length === 0) return null;
              return (
                <div key={grupo.tipo} style={{ marginBottom: 24 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>{grupo.label} ({lista.length})</div>
                  <div style={{ fontSize: 12, color: "#c0c0e0", marginBottom: 10 }}>{grupo.desc}</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: 8 }}>
                    {lista.map((p) => {
                      const dias = diasParaVencer(p.periodo);
                      return (
                        <div key={p.id} onClick={() => setPropSel(p)} style={{ background: "#252538", border: "1px solid #3d3b6e", borderRadius: 8, padding: "12px 14px", cursor: "pointer" }}>
                          <div style={{ fontWeight: 500, fontSize: 13, marginBottom: 3 }}>{p.direccion}</div>
                          <div style={{ fontSize: 12, color: "#d0d0f0", marginBottom: 6 }}>{p.inquilino || "Sin inquilino"} {p.telefono ? `· ${p.telefono}` : ""}</div>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: 11, color: "#c0c0e0" }}>{p.periodo || "Sin período"}</span>
                            {dias !== null && <Badge tipo={semaforo(p)} />}
                          </div>
                          {p.mail && <div style={{ fontSize: 11, color: "#89b4fa", marginTop: 4 }}>{p.mail}</div>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            {stats.vencidos + stats.proximos + stats.libres === 0 && (
              <div style={{ background: "#2d2b55", borderRadius: 10, padding: 40, textAlign: "center", color: "#7c6fcd" }}>
                <div style={{ fontSize: 15, fontWeight: 600 }}>Todo en orden</div>
                <div style={{ fontSize: 13, color: "#d0d0f0", marginTop: 4 }}>No hay contratos vencidos ni propiedades sin gestionar.</div>
              </div>
            )}
          </div>
        )}
      </div>

      {propSel && <ModalDetalle prop={propSel} onClose={() => setPropSel(null)} onSave={guardarProp} onDelete={eliminarProp} />}
      {cobroModal && <ModalCobro prop={cobroModal} onClose={() => setCobroModal(null)} onGuardar={guardarCobro} />}
      {nuevaModal && <ModalNueva onClose={() => setNuevaModal(false)} onGuardar={guardarNueva} />}
    </div>
  );
}
