VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSummary 
      Caption         =   "Summary"
      Height          =   2535
      Left            =   360
      TabIndex        =   19
      Top             =   5040
      Width           =   7335
      Begin VB.TextBox txtAverageDiscount 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtDiscountedAmountSum 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtDiscountSum 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtQuantitySum 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblAverageDiscount 
         Caption         =   "Average Discount"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblDiscountedAmountSum 
         Caption         =   "Total of Discounted Amounts"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblDiscountSum 
         Caption         =   "Total of Discounts Given"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label lblQuantitySum 
         Caption         =   "Total Number of Books"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6360
      TabIndex        =   18
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   615
      Left            =   6360
      TabIndex        =   17
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearSale 
      Caption         =   "Clear Sale"
      Height          =   615
      Left            =   4800
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   4800
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   4215
      Begin VB.TextBox txtExtendedPrice 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtDiscountedPrice 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtDiscount 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblDiscountedPrice 
         Caption         =   "Discounted Price"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDiscount 
         Caption         =   "15% Discount"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblExtendedPrice 
         Caption         =   "Extended Price"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   7455
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblPrice 
         Caption         =   "Price"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblBookSales 
      Alignment       =   2  'Center
      Caption         =   "Book Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dimension module-level variables

Dim mintQuantitySum         As Integer
Dim mcurDiscountSum         As Currency
Dim mcurDiscountedPriceSum  As Currency
Dim mintSaleCount           As Integer
Dim mcurAverageDiscount     As Currency
Const mcurDiscountRate      As Currency = 0.15

Private Sub cmdCalculate_Click()
    'Calculate the price and discount
    
    Dim intQuantity         As Integer
    Dim curPrice            As Currency
    Dim curExtendedPrice    As Currency
    Dim curDiscount         As Currency
    Dim curDiscountedPrice  As Currency
    
    'Convert input values to numeric variables
    intQuantity = Val(txtQuantity.Text)
    curPrice = Val(txtPrice.Text)
    
    'Calculate values for sale
    curExtendedPrice = intQuantity * curPrice
    curDiscount = curExtendedPrice * mcurDiscountRate
    curDiscountedPrice = curExtendedPrice - curDiscount
    
    'Calculate values for summary
    mintQuantitySum = mintQuantitySum + intQuantity
    mcurDiscountSum = mcurDiscountSum + curDiscount
    mcurDiscountedPriceSum = mcurDiscountedPriceSum + curDiscountedPrice
    mintSaleCount = mintSaleCount + 1
    mcurAverageDiscount = mcurDiscountSum / mintSaleCount
    
    'Format and display answers for sale
    txtExtendedPrice.Text = FormatCurrency(curExtendedPrice)
    txtDiscount.Text = FormatCurrency(curDiscount, 2)
    txtDiscountedPrice.Text = FormatCurrency(curDiscountedPrice)
    
    'Format and display answers for summary
    txtQuantitySum.Text = mintQuantitySum
    txtDiscountSum.Text = FormatCurrency(mcurDiscountSum)
    txtDiscountedAmountSum.Text = FormatCurrency(mcurDiscountedPriceSum)
    txtAverageDiscount.Text = FormatCurrency(mcurAverageDiscount)
    
End Sub

Private Sub cmdClearSale_Click()
    'Clear previous amounts from the form
    txtQuantity.Text = ""
    txtTitle.Text = ""
    txtPrice.Text = ""
    txtExtendedPrice.Text = ""
    txtDiscount.Text = ""
    txtDiscountedPrice.Text = ""
    txtQuantity.SetFocus
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    PrintForm
End Sub
