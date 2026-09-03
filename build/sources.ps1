# The product's C# sources, in compile order. Dot-source this from every script
# that compiles the product (the packer, the compile check and the tests), so
# the list lives in exactly one place.
$RdvSources = @(
  'Rdv3Core.cs', 'Rdv3Index.cs', 'Rdv3Ledger.cs', 'Rdv3Xlsx.cs', 'Rdv3Shared.cs', 'Rdv3Jobs.cs',
  'Rdv3Json.cs', 'Rdv3Data.cs', 'Rdv3Process.cs', 'Rdv3Config.cs', 'Rdv3Screen.cs', 'Rdv3Values.cs',
  'Rdv3Uia.cs', 'Rdv3Watch.cs', 'Rdv3Text.cs',
  'Rdv3Ui.cs', 'Rdv3Modals.cs', 'Rdv3Settings.cs', 'Rdv3App.cs')
