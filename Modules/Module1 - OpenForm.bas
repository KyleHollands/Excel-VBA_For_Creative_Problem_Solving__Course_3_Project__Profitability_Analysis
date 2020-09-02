Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub OpenForm()

MainForm.costOfLandChance1.Text = "20"
MainForm.costOfLandChance2.Text = "50"
MainForm.costOfLandChance3.Text = "30"
MainForm.costOfLandCost1.Text = "-3.0"
MainForm.costOfLandCost2.Text = "-5.0"
MainForm.costOfLandCost3.Text = "-10.0"

MainForm.costOfRoyaltiesLow.Text = "-1.5"
MainForm.costOfRoyaltiesMode.Text = "-4.0"
MainForm.costOfRoyaltiesHigh.Text = "-5.0"

MainForm.totalDepCapitalAve.Text = "-100"
MainForm.totalDepCapitalStDev.Text = "20"

MainForm.workingCapitalMin.Text = "-20"
MainForm.workingCapitalMax.Text = "-40"

MainForm.startupCostsAve.Text = "-12.5"
MainForm.startupCostsStDev.Text = "3"

MainForm.salesRevenueLow.Text = "30"
MainForm.salesRevenueMode.Text = "50"
MainForm.salesRevenueHigh.Text = "60"

MainForm.prodCostsLow.Text = "-5"
MainForm.prodCostsMode.Text = "-6"
MainForm.prodCostsHigh.Text = "-8"

MainForm.taxChance1.Text = "30"
MainForm.taxChance2.Text = "70"
MainForm.taxRate1.Text = "0.35"
MainForm.taxRate2.Text = "0.40"

MainForm.interestRateMin.Text = "0.09"
MainForm.interestRateMax.Text = "0.15"

MainForm.numOfSimulations.Text = "1000"

MainForm.Show

End Sub
