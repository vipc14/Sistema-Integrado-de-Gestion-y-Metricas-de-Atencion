Sistema Integrado de Gestión y Métricas de Atención
es una aplicación web completa desarrollada sobre la plataforma de Google Apps Script y Google Sheets para el registro, seguimiento y visualización de indicadores clave de rendimiento (KPIs) en centros de atención al cliente.

📜 Descripción
Este sistema proporciona una solución centralizada para que los supervisores registren métricas de rendimiento semanales a través de diferentes canales de atención (Presencial, Telefónico y Virtual). Los datos se almacenan en una hoja de cálculo de Google, y la aplicación web ofrece paneles de control interactivos para visualizar y analizar esta información, facilitando la toma de decisiones basada en datos.

✨ Características Principales
Panel de Control Interactivo: Visualiza un resumen de los KPIs más importantes con filtros dinámicos por canal, subcanal (proveedor/OSAC), semana o rangos de fecha personalizados.

Comparativa de Canales: Analiza el rendimiento histórico y actual de cada canal de atención (Presencial, Telefónico, Virtual) en una vista consolidada.

Gráficos Dinámicos: Gráficos de barras generados con Chart.js para mostrar la evolución de las atenciones a lo largo del tiempo.

Registro de Datos Sencillo: Formularios intuitivos y separados por canal para que los supervisores ingresen los indicadores de rendimiento de forma rápida y precisa.

Gestión de Usuarios Basada en Roles:

Administrador: Control total sobre la gestión de usuarios (crear, editar, activar/desactivar, eliminar) y la configuración del sistema.

Supervisor: Permisos para registrar nuevos datos y visualizar todos los paneles.

Usuario: Rol de solo lectura para visualizar los paneles de control.

Sistema de Autenticación Seguro: Inicio de sesión con contraseñas encriptadas (SHA-256 con sal) para proteger el acceso.

Caché de Servidor: Utiliza el CacheService de Google Apps Script para acelerar la carga de datos recurrentes como la lista de semanas y subcanales.

Interfaz Moderna: Diseño responsivo con tema claro y oscuro para una mejor experiencia de usuario.

🚀 Arquitectura y Tecnologías
El sistema está construido íntegramente sobre el ecosistema de Google, lo que lo hace robusto, escalable y sin costo de hosting.

Backend: Google Apps Script (Code.gs)

Maneja toda la lógica de negocio, incluyendo la autenticación, el acceso a datos, los cálculos de KPIs y la gestión de usuarios.

Actúa como una API que se comunica con el frontend.

Base de Datos: Google Sheets

Almacena todos los datos de la aplicación en diferentes hojas: Data, Usuarios, Config y Semanas.

Frontend: HTML, CSS y JavaScript Vainilla (Index.html, Css.html, Js.html)

Construye la interfaz de usuario que se muestra al usuario final.

Se comunica con el backend de Apps Script de forma asíncrona usando google.script.run.

Utiliza Chart.js para la visualización de datos.



