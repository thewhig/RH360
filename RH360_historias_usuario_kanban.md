# RH360 - Historias de usuario y tablero Kanban

## Datos del proyecto

- Nombre de la aplicación: RH360
- Código de la aplicación: RH3
- Sistema: RRHH. Recursos Humanos
- Subsistema: GPRH. Gestión de Personal
- Nombre del tablero Kanban: RH360

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

## Instrucciones para crear el tablero manualmente en GitHub

1. Acceder a GitHub con la cuenta que se utilizará para la entrega.
2. Entrar en el perfil de usuario y seleccionar Projects.
3. Crear un nuevo proyecto con nombre RH360.
4. Seleccionar vista de tipo Board o Kanban.
5. Configurar la visibilidad como pública, o garantizar permiso de lectura al profesor.
6. Editar el campo Status para que contenga exactamente estos valores: Backlog del producto, Pendiente de desarrollo, En proceso de desarrollo, En fase de pruebas y estrés, Pendiente de despliegue, Finalizada / Histórico.
7. Crear las historias indicadas en este documento y asignar cada una al estado correspondiente.
8. Copiar la URL pública del proyecto y añadirla en la entrega del aula virtual.

## Comandos orientativos con GitHub CLI

Estos comandos requieren tener instalado GitHub CLI (`gh`) y haber iniciado sesión con una cuenta con permisos para crear repositorios y proyectos.

```powershell
gh auth login
gh repo create RH360 --public --description "Aplicación corporativa para la gestión de comunicaciones internas, diarios, agenda de empleados y trámites del empleado." --add-readme
gh project create --owner "@me" --title "RH360"
```

Después de crear el proyecto, se deben configurar manualmente los estados del campo Status si la versión instalada de GitHub CLI no permite modificar opciones de campos de Projects v2 directamente.

Referencia oficial consultable: https://docs.github.com/issues/planning-and-tracking-with-projects/automating-your-project/using-the-api-to-manage-projects
