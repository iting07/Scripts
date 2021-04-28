resource "azurerm_app_service" "app" {
  name                = "wapp-${var.project}-${var.environment}"
  location            = azurerm_resource_group.rg.location
  resource_group_name = azurerm_resource_group.rg.name
  app_service_plan_id = azurerm_app_service_plan.asp.id

  tags = local.common_tags

  https_only = true
}
