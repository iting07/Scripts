locals {
  # Common tags to be assigned to all resources
  common_tags = {
    Environment      = upper(var.environment)
    TeamName         = "PlatEng"
    Tribe            = "Enable"
    LastModifiedDate = timestamp()
  }
}
