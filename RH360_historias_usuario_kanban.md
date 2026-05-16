# RH360 - Historias de usuario y tablero Kanban

## Datos del proyecto

- Nombre de la aplicación: RH360
- Código de la aplicación: RH3
- Sistema: RRHH. Recursos Humanos
- Subsistema: GPRH. Gestión de Personal
- Nombre del tablero Kanban: RH360
- Repositorio creado: https://github.com/thewhig/RH360
- Issues creadas: https://github.com/thewhig/RH360/issues
- Tablero Kanban público: https://github.com/users/thewhig/projects/3
- Estado del GitHub Project v2: configurado.

## Estados obligatorios

- Backlog del producto
- Pendiente de desarrollo
- En proceso de desarrollo
- En fase de pruebas y estrés
- Pendiente de despliegue
- Finalizada / Histórico

## Distribución propuesta de historias

### Finalizada / Histórico

- Definir nombre, sistema, subsistema y código de la aplicación.
- Analizar las necesidades trasladadas por Recursos Humanos.

### Pendiente de despliegue

- Como responsable de Recursos Humanos, quiero dar de alta empleados, para permitirles acceder a la aplicación.

### En fase de pruebas y estrés

- Como empleado, quiero consultar la agenda de compañeros, para localizar a personas de otros turnos o áreas.
- Como empleado, quiero marcar comunicaciones como revisadas, para controlar qué información ya he atendido.

### En proceso de desarrollo

- Como empleado, quiero crear diarios por temática, para organizar mis anotaciones laborales.
- Como empleado, quiero enviar comunicaciones escritas a otros compañeros, para informarles aunque no coincidamos en el mismo turno.
- Como administrador, quiero configurar departamentos, áreas y turnos, para organizar correctamente la información de la empresa.

### Pendiente de desarrollo

- Como empleado, quiero buscar compañeros por nombre, departamento o turno, para encontrar rápidamente a la persona adecuada.
- Como empleado, quiero registrar notas e incidencias, para dejar constancia de información importante.
- Como empleado, quiero consultar las comunicaciones recibidas, para revisar avisos o incidencias pendientes.
- Como responsable de Recursos Humanos, quiero modificar los datos de empleados, para mantener actualizada la información interna.
- Como responsable de Recursos Humanos, quiero dar de baja empleados, para impedir el acceso a personas que ya no pertenezcan a la organización.
- Como administrador, quiero gestionar roles y permisos, para controlar qué puede hacer cada tipo de usuario.
- Como responsable o mando intermedio, quiero consultar comunicaciones relevantes de mi área, para hacer seguimiento de incidencias importantes.
- Como empleado, quiero consultar el histórico de mis notas y comunicaciones, para recuperar información registrada anteriormente.

### Backlog del producto

- Como empleado, quiero solicitar una excedencia desde la aplicación, para evitar entregar formularios en papel.
- Como responsable de Recursos Humanos, quiero aprobar o rechazar solicitudes de excedencia, para agilizar la tramitación.
- Como empleado, quiero consultar el histórico de mis solicitudes de excedencia, para conocer el estado y resultado de mis trámites.
- Como responsable de Recursos Humanos, quiero consultar el histórico de trámites de un empleado, para revisar sus solicitudes anteriores.

## Configuración realizada en GitHub

1. Proyecto GitHub configurado con nombre RH360.
2. Visibilidad del proyecto configurada como pública.
3. Proyecto vinculado al repositorio público RH360.
4. Campo Status configurado con los seis estados solicitados.
5. Historias de usuario añadidas como issues y asignadas al estado correspondiente.

Distribución verificada:

- Backlog del producto: 4
- Pendiente de desarrollo: 8
- En proceso de desarrollo: 3
- En fase de pruebas y estrés: 2
- Pendiente de despliegue: 1
- Finalizada / Histórico: 2

## Comandos usados con GitHub CLI

Estos comandos requieren tener instalado GitHub CLI (`gh`) y haber iniciado sesión con una cuenta con permisos para editar GitHub Projects.

```powershell
.\tools\gh\bin\gh.exe auth login --hostname github.com --git-protocol https --web --scopes "project,repo,read:org"
.\tools\gh\bin\gh.exe project edit 3 --owner thewhig --title RH360 --visibility PUBLIC
.\tools\gh\bin\gh.exe project link 3 --owner thewhig --repo RH360
```

Referencia oficial consultable: https://docs.github.com/issues/planning-and-tracking-with-projects/automating-your-project/using-the-api-to-manage-projects
