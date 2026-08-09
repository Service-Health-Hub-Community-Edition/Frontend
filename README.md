# Service Health Hub Web App

Service Health Hub Web App is the browser and Microsoft Teams user interface for Service Health Hub. It provides dashboards, reporting, administration, and integration configuration for Microsoft cloud service health, change communications, roadmap updates, endpoint changes, license insights, and organization-specific announcements.

The project consists of a .NET API backend and a React / TypeScript frontend.

## Overview

Service Health Hub helps IT operations, service owners, change managers, and support teams centralize Microsoft cloud communications and operationalize them through dashboards, notifications, task systems, and custom integration points.

The Web App focuses on interactive user and administrator experiences. Background synchronization and scheduled processing can be handled by connected services such as Azure Functions, Logic Apps, Power Automate flows, or other configured integration endpoints.

## Key Capabilities

### Dashboards and Reporting

The application provides centralized views for:

* Active incidents and advisories
* Upcoming Microsoft cloud changes
* Historical service health information
* Message Center and roadmap communications
* Microsoft 365 endpoint changes
* License statistics and forecasts
* Organization-specific service announcements

These views are intended to help teams understand current service impact, prepare for upcoming changes, review historical incidents, and identify operational follow-up actions.

### Integration Triggering

Service Health Hub can trigger configured integration points so customers can bring their own business logic into the service health and change management process.

Supported integration patterns may include:

* Power Automate flows
* Azure Logic Apps
* Azure Functions
* Webhook-based custom actions
* Queue-based or event-based integrations
* Customer-specific automation endpoints

Integration points can be triggered:

* Manually by a user from the Web App
* Automatically through notification routing rules
* Per data source, depending on configuration
* As part of custom operational workflows

### Admin Center

The Web App includes an Admin Center for configuration and operational administration.

Administrators can configure:

* Notification connectors and routing per data source
* Task field mapping for DevOps, ServiceNow, or Jira
* Copilot Connector configuration
* Integration points, also known as custom actions, per data source
* System configuration
* Language service integration
* Translator service integration

## Functional Areas

### User Experience

The user-facing area provides access to synchronized service communications and operational insights.

Typical user scenarios include:

* Review active Microsoft 365, Azure, Dynamics 365, and Power Platform incidents
* Track upcoming changes from Microsoft communications
* Review roadmap and Message Center items
* Search or filter historical service health information
* Open related work items or external task records
* Trigger configured custom actions where permitted

### Administration Experience

The Admin Center provides the configuration interface for operational owners.

Typical administrator scenarios include:

* Configure notification routing rules
* Define connector targets for each data source
* Map source metadata to task system fields
* Configure custom integration points
* Enable or disable data source-specific behavior
* Configure language and translation services
* Manage Copilot Connector settings

## Technology Stack

### Backend

* .NET API
* REST API endpoints
* Authentication and authorization integration
* Configuration and administration services
* Integration orchestration layer
* Data access layer for operational and configuration data

### Frontend

* React
* TypeScript
* Component-based UI architecture
* Dashboard and reporting views
* Admin Center configuration pages
* Browser access and Microsoft Teams-friendly user experience

### Integrations

The solution is designed to integrate with customer-specific tools and automation platforms, including:

* Azure DevOps
* ServiceNow
* Jira
* Power Automate
* Azure Logic Apps
* Azure Functions
* Microsoft Teams notification endpoints
* Copilot connector configuration
* Azure Language services
* Azure Translator services

## Data Sources

Service Health Hub is designed around Microsoft cloud communications and operational data sources such as:

* Microsoft 365 Service Health
* Microsoft 365 Message Center
* Microsoft 365 Roadmap
* Microsoft 365 endpoint changes
* Azure Service Health
* Azure Updates
* Dynamics 365 release information
* Power Platform release information
* Organization-specific service announcements
* License statistics and forecast data

## Configuration

### Notification Connectors and Routing

Notification routing is configured per data source. Routing rules can determine where communications are sent based on metadata such as service, title, category, impact, or other source-specific properties.

Connector targets may include:

* Microsoft Teams channels
* Email destinations
* Service Bus queues
* Custom integration endpoints

### Task Field Mapping

Task field mapping controls how source metadata is written to external task systems.

Supported task system targets may include:

* Azure DevOps work items
* ServiceNow records
* Jira issues

Mappings should define:

* Source entity property
* Destination field
* Whether the value is used during create operations
* Whether the value is used during update operations
* Optional transformation or formatting rules

### Integration Points and Custom Actions

Integration points allow customers to attach their own automation to Service Health Hub events.

Examples:

* Trigger a Power Automate flow for high-impact incidents
* Trigger a Logic App for selected roadmap communications
* Invoke an Azure Function for organization-specific enrichment
* Send a payload to an external ITSM or CMDB process

### Language and Translator Services

The system configuration can include Azure Language and Azure Translator services to support classification, enrichment, translation, or localized notifications.

## Security and Access Control

Recommended security principles:

* Use Microsoft Entra ID for authentication
* Use role-based access control for user and administrator features
* Apply least privilege for service principals and managed identities
* Store secrets in Azure Key Vault or an equivalent secure secret store
* Avoid storing secrets in source control
* Restrict administrative configuration to authorized groups only
* Review integration endpoints before enabling automatic triggers

