resource "azurerm_app_service_plan" "asp" {
  name                = "asp-${var.project}-${var.environment}"
  resource_group_name = azurerm_resource_group.rg.name
  location            = azurerm_resource_group.rg.location
  kind                = "Linux"
  reserved            = true

  tags = local.common_tags

  sku {
    tier = "Basic"
    size = "B1"
  }
}
