# LeadMailer — Automatización de correos para leads desde Excel

Aplicación de escritorio **WPF (.NET 8 / Windows)** para gestionar y enviar correos automáticos a leads cargados desde un Excel diario.

---

## 🚀 Abrir en Visual Studio

1. Descomprime la carpeta `LeadMailer_VS`
2. Abre **`LeadMailer.sln`** con Visual Studio 2022
3. Visual Studio restaurará los paquetes NuGet automáticamente
4. Pulsa **F5** para compilar y ejecutar

### Requisitos
- Visual Studio 2022 (Community, Professional o Enterprise)
- .NET 8 SDK (incluido en VS 2022 v17.8+)
- Workload: **.NET Desktop Development**

---

## ⚙️ Configuración inicial (primer uso)

### 1. Configurar SMTP
Ve a la pestaña **⚙ Configuración** y rellena:

| Campo | Gmail |
|-------|-------|
| Host | `smtp.gmail.com` |
| Puerto | `587` |
| TLS | ✓ Activado |
| Usuario | `tucorreo@gmail.com` |
| Contraseña | Contraseña de aplicación |
| Nombre remitente | Nombre que verá el destinatario |
| Email remitente | tucorreo@gmail.com |

> **Gmail** → Seguridad → Verificación en 2 pasos → **Contraseñas de aplicación**

### 2. Cargar el Excel diario
1. Ve a **📋 Leads**
2. Clic en **📂 Examinar** → selecciona el `.xlsx`
3. Clic en **🔄 Cargar leads**
4. Si aparece aviso amarillo → hay cursos nuevos sin configurar

### 3. Configurar cursos nuevos
1. Ve a **🎓 Cursos**
2. Las tarjetas con ⚠ no están configuradas
3. Clic en **✏ Editar** y rellena:
   - Nombre corto, asunto del email
   - Texto de marketing
   - Fechas inicio/fin y horario
   - Requisitos de acceso
   - Documentación necesaria
   - URL ficha de inscripción

### 4. Enviar correos
1. Vuelve a **📋 Leads**
2. Los leads nuevos ya están seleccionados (✓)
3. Los ya enviados aparecen marcados como **Enviado** y no se pueden seleccionar
4. Clic en **🚀 Enviar correos seleccionados**

---

## 📊 Formato del Excel

El archivo Excel debe tener estas columnas en la primera fila (el orden no importa):

| Columna | Descripción |
|---------|-------------|
| `Curso` | Identificador del curso (detecta cursos nuevos automáticamente) |
| `Nombre` | Nombre completo del lead |
| `Email` | Dirección de correo |
| `Teléfono` | Teléfono de contacto |
| `Provincia` | Provincia del lead |
| `Plataforma` | Origen (fb, ig, etc.) |
| `Situación Laboral` | Situación laboral del lead |
| `Sector Laboral / Nivel Estudios` | Nivel formativo |
| `Aceptación Publicidad` | true/false |
| `OBSERVACIONES` | Notas internas |

---

## 💾 Almacenamiento de datos

Los datos se guardan automáticamente en:
```
C:\Users\<usuario>\AppData\Roaming\LeadMailer\data.json
```
Contiene: configuración SMTP, información de cursos y registro de envíos.

---

## 🏗️ Estructura del proyecto

```
LeadMailer.sln
LeadMailer/
├── App.xaml / App.xaml.cs
├── LeadMailer.csproj
├── Models/
│   └── Models.cs               → Lead, CourseInfo, SmtpConfig, AppData, SentRecord
├── Services/
│   ├── DataService.cs          → Persistencia JSON
│   ├── ExcelService.cs         → Lectura Excel (EPPlus)
│   └── EmailService.cs         → Envío SMTP (MailKit)
├── ViewModels/
│   └── MainViewModel.cs        → Toda la lógica (MVVM + CommunityToolkit)
├── Views/
│   ├── MainWindow.xaml         → UI completa
│   └── MainWindow.xaml.cs      → Code-behind (PasswordBox)
├── Converters/
│   └── Converters.cs           → Value converters WPF
└── Themes/
    ├── Colors.xaml             → Paleta de colores
    └── Controls.xaml           → Estilos de controles
```

---

## 📦 Paquetes NuGet

| Paquete | Versión | Uso |
|---------|---------|-----|
| EPPlus | 7.3.2 | Lectura de Excel |
| MailKit | 4.8.0 | Envío SMTP moderno |
| MimeKit | 4.8.0 | Construcción de emails HTML |
| Newtonsoft.Json | 13.0.3 | Serialización de datos |
| CommunityToolkit.Mvvm | 8.3.2 | MVVM con source generators |
| System.Web.HttpUtility | 7.0.0 | HtmlEncode para el correo HTML |
