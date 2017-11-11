// VisualShow3DLightObject.cs
// Librairie VisualShow3DLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;


namespace VisualShow3DLight
  {
  class VisualShow3DLightUtil
    {

    internal class NativeMethods
      {
      /// <summary>Windows constant - Sets a new window style.</summary>
      internal const short GWL_STYLE = (-16);
      /// <summary>Windows constant - Creates a child window..</summary>
      internal const int WS_CHILD = 0x40000000;
      /// <summary>Windows constant - Creates a window that is initially
      /// visible.</summary>
      internal const int WS_VISIBLE = 0x10000000;
      /// <summary>Windows constant - Creates a window that is initially
      /// visible.</summary>
      internal const int WM_CLOSE = 0x0010;
      /// <summary>Declare a private constructor to prevent new instances
      /// of the NativeMethods class from being created. This constructor
      /// is intentionally left blank.</summary>
      /// 
      internal const int SW_HIDE = 0;
      internal const int SW_SHOWNORMAL = 1;
      internal const int SW_SHOWNOACTIVATE = 4;
      internal const int SW_SHOW = 5;
      internal const int SW_MINIMIZE = 6;
      internal const int SW_SHOWNA = 8;
      internal const int SW_SHOWMAXIMIZED = 11;
      internal const int SW_MAXIMIZE = 12;
      internal const int SW_RESTORE = 13;
      internal const int WM_KEYDOWN = 0x0100;
      internal const int WM_CHAR = 0x0102;
      internal const int VK_DOWN = 0x0028;
      internal const int VK_UP = 0x0026;
      internal const int VK_PRIOR = 0x0021;
      internal const int VK_NEXT = 0x0022;
      internal const int VK_LEFT = 0x0025;
      internal const int VK_RIGHT = 0x0027;
      internal const uint KF_EXTENDED = 0x00010000;
      internal const uint KF_ALTDOWN = 0x20000000;
      





      private NativeMethods()
        {
        // No initialization is required.
        }

      [StructLayout(LayoutKind.Sequential)]
      public struct RECT
        {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
        }
      /// <summary>
      /// Prototype de SetParent() pour PInvoke</summary>
      /// </summary>
      /// <param name="hWndChild"></param>
      /// <param name="hWndNewParent"></param>
      /// <returns></returns>
      [System.Runtime.InteropServices.DllImport("user32.dll")]
      internal static extern int SetParent(int hWndChild,
        int hWndNewParent);
      /// <summary>Prototype of SetWindowLong() for PInvoke</summary>
      [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
      internal static extern int SetWindowLongW(int hwnd,
        int nIndex,
        int dwNewLong);
      // GetParent
      [System.Runtime.InteropServices.DllImport("user32.dll")]
      internal static extern int GetParent(int hWnd);
      // Send message
      [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
      internal static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
      }


    public static string AddQuotes(string inStr)
      {
      //Note:  Lors de la construction d'une chaîne, il faut
      // s'assurer qu'il n'y a pas de """ deadnas sinon, il faut les remplacer
      // par des simple quotes "'"
      const string quoteStr = "\"";

      string outStr = quoteStr + inStr + quoteStr;
      return outStr;
      }

    public static string StripQuotes(string inStr)
      {
      char[] trimChr = { '"' };
      string outStr = "";

      if (inStr != null)
        {
        outStr = inStr.TrimStart(trimChr);
        outStr = outStr.TrimEnd(trimChr);
        }
      return outStr;
      }

    public static bool GetActiveSelection(Visio.Application visApp, out Visio.Selection visSelection)
      {
      Visio.Window visWindowDoc = null;

      visSelection = null;
      try
        {
        //Récupère la fenêtre courante
        if (GetVisActiveWindow(visApp, ref visWindowDoc))
          {
          //Récupère la sélection courante
          visSelection = visWindowDoc.Selection;
          }
        }
      catch
        {
        return false;
        }
      return true;
      }

    public static bool GetVisActiveWindow(Visio.Application visApp, ref Visio.Window visWindowDoc)
      {
      try
        {
        visWindowDoc = visApp.ActiveWindow;
        }
      catch
        {
        return false;
        }
      return true;
      }

    public static bool IsOneD(Visio.Shape visShape)
      {

      try
        {
        if (visShape != null)
          {
          if (visShape.OneD == -1)
            return true;
          else
            return false;
          }
        return false;
        }
      catch
        {
        return false;
        }
      }

    public static bool IsSectionExist(Visio.Shape visShape, short iSection,
                                      out Visio.Section visSection)
      {
      bool bExist;
      short fexist = 0;

      if (visShape.get_SectionExists(iSection, fexist) != 0)
        {
        visSection = visShape.get_Section(iSection);
        bExist = true;
        }
      else
        {
        visSection = null;
        bExist = false;
        }
      return bExist;
      }

    public static bool GetRowIndex(Visio.Shape visShape, int visSectionIndice,
                                   string strNameU, out int iRowIndex)
      {
      Visio.Section visSection;
      Visio.Row visRow;
      int iNbRow;
      string strLocNameU;

      iRowIndex = -1;
      try
        {
        visSection = visShape.get_Section((short)visSectionIndice);
        iNbRow = visSection.Count;
        for (int i = 0; i < iNbRow; i++)
          {
          visRow = visSection[(short)i];
          strLocNameU = GetNameU(visRow);
          if (strLocNameU == strNameU)
            {
            iRowIndex = i;
            return true;
            }
          }
        }
      catch
        {
        return false;
        }
      return false;
      }

    public static bool AddShapeDataRow(Visio.Shape visShape, string strRowName, string strLabel, int iType, string strFormat,
                                   string strValue, int iLanguage)
      {
      int iLocVisSectionIndice, iLocRow, iLocRowControl;
      Visio.Section visSection;
      Visio.Row visRow;
      Visio.Cell visCell;
      bool bReturn = false;


      if (!IsSectionExist(visShape, (int)Visio.VisSectionIndices.visSectionProp, out visSection))
        {
        iLocVisSectionIndice = visShape.AddSection((int)Visio.VisSectionIndices.visSectionProp);
        }
      if (!GetRowIndex(visShape, (int)Visio.VisSectionIndices.visSectionProp, strRowName, out iLocRow))
        {
        iLocRow = visShape.AddRow((int)Visio.VisSectionIndices.visSectionProp, (int)Visio.VisRowIndices.visRowLast,
                                              (int)Visio.VisRowTags.visTagDefault);
        }
      // Vérification de l'existence de la ligne de nom strNameU
      // En effet dans certain cas la ligne est déja présente
      // car dans ces cas le delete section cache la section 
      // sans supprimer les lignes, il faut donc vérifier que la ligne
      // n'existe pas avant de la nommer
      if (!GetRowIndex(visShape, (int)Visio.VisSectionIndices.visSectionProp, strRowName, out iLocRowControl))
        {
        // on la nomme
        visSection = visShape.get_Section((int)Visio.VisSectionIndices.visSectionProp);
        visRow = visSection[(short)iLocRow];
        visRow.Name = strRowName;
        // On met le libellé
        visCell = visRow.get_CellU((short)Visio.VisCellIndices.visCustPropsLabel);
        SetStringCellVal(visCell, strLabel);
        // On met le type
        visCell = visRow.get_CellU((short)Visio.VisCellIndices.visCustPropsType);
        SetIntCellVal(visCell, iType);
        // On met le format
        visCell = visRow.get_CellU((short)Visio.VisCellIndices.visCustPropsFormat);
        SetStringCellVal(visCell, strFormat);
        // On met la valeur
        visCell = visRow.get_CellU((short)Visio.VisCellIndices.visCustPropsValue);
        SetStringCellVal(visCell, strValue);
        // On met le langage
        visCell = visRow.get_CellU((short)Visio.VisCellIndices.visCustPropsLangID);
        SetIntCellVal(visCell, iLanguage);
        }
      bReturn = true;
      return bReturn;
      }


        public static string GetNameU(Visio.Row visRow)
      {

      return visRow.NameU;
      }

    public static bool GetIntCellVal(Visio.Cell visCell, int visUnits, out int iValue)
      {

      iValue = visCell.get_ResultInt(visUnits, 1);
      return true;
      }
    public static bool GetIntCellVal(Visio.Page visPage, string strCellName, out int iValue)
      {
      Visio.Shape visShapePage;
      Visio.Cell visCell;
      iValue = -1;

      try
        {
        visShapePage = visPage.PageSheet;
        visCell = visShapePage.get_CellsU(strCellName);
        return GetIntCellVal(visCell, (int)Visio.VisUnitCodes.visNumber, out iValue);
        }
      catch
        {
        return false;
        }
      }

    public static bool GetDoubleCellVal(Visio.Cell visCell, int visUnits, out double dblValue)
      {

      dblValue = visCell.get_Result(visUnits);
      return true;
      }

    public static bool GetDoubleCellVal(Visio.Shape visShape, string strCellName, int visUnits, out double dblValue)
      {
      Visio.Cell visCell;

      visCell = visShape.get_CellsU(strCellName);
      return GetDoubleCellVal(visCell, visUnits, out dblValue);
      }

    public static bool GetDoubleCellVal(Visio.Page visPage, string strCellName, int visUnits, out double dblValue)
      {
      Visio.Shape visShapePage;
      Visio.Cell visCell;
      dblValue = -1.0;

      try
        {
        visShapePage = visPage.PageSheet;
        visCell = visShapePage.get_CellsU(strCellName);
        return GetDoubleCellVal(visCell, visUnits, out dblValue);
        }
      catch
        {
        return false;
        }
      }

    public static bool GetStringCellFormula(Visio.Cell visCell, out string strChaine)
      {

      strChaine = visCell.FormulaU;
      if (strChaine != String.Empty)
        {
        int longueur;

        if ((longueur = strChaine.Length) >= 2)
          strChaine = strChaine.Substring(1, longueur - 2);
        else
          strChaine = String.Empty;
        }
      return true;
      }

    public static bool GetStringCellProp(Visio.Shape visShape, string strFullName, out string strValue)
      {
      Visio.Cell visCell;

      if (visShape.get_CellExists(strFullName, (short)Visio.VisExistsFlags.visExistsAnywhere) != 0)
        {
        // La cellule existe
        visCell = visShape.get_CellsU(strFullName);
        GetStringCellFormula(visCell, out strValue);
        return true;
        }
      else
        {
        strValue = "";
        return false;
        }
      }

    public static bool SetStringCellVal(Visio.Cell visCell, string strChaine)
      {

      visCell.FormulaForceU = AddQuotes(StripQuotes(strChaine));
      return true;
      }

    public static bool SetIntCellVal(Visio.Cell visCell, int iValue)
      {

      try
        {
        visCell.set_ResultFromIntForce((int)Visio.VisUnitCodes.visNumber, iValue);
        }
      catch (System.Exception except)
        {
        string strMessage = except.Message;
        return false;
        }
      return true;
      }

    public static bool GetFormulaUCell(Visio.Shape visShape, string strCellName, out string strFormula)
      {
      Visio.Cell visCell;

      try
        {
        visCell = visShape.get_CellsU(strCellName);
        return GetFormulaUCell(visCell, out strFormula);
        }
      catch
        {
        strFormula = "";
        return false;
        }
      }

    public static bool GetFormulaUCell(Visio.Cell visCell, out string strFormula)
      {

      strFormula = visCell.FormulaU;
      return true;
      }

    public static bool GetRGBCellVal(Visio.Shape visShape, string strCellName, out string strValue)
      {
      strValue = "";
      Visio.Cell visCell;
      try
        {
        visCell = visShape.get_CellsU(strCellName);
        strValue = visCell.get_ResultStr((int)Visio.VisUnitCodes.visUnitsColor);
        }
      catch
        {
        }
      return true;
      }

    public static bool AddAnchorWindowToVisio(Visio.Application visApplication, string strCaption,
                                              int iAnchorMode, bool bAutoHide, bool bVisible,
                                              int iLeft, int iTop, int iWidth, int iHeight,
                                              string strMergeID, string strMergeClass, int iMergePosition,
                                              ref Visio.Window visWindowToAdd)
      {
      Visio.Window visActiveWindow, visNewWindow;
      bool bSuccess = false;
      int iWindowState, iWindowType;

      visActiveWindow = visApplication.ActiveWindow;
      if (visActiveWindow != null)
        {
        iWindowState = iAnchorMode;
        if (bAutoHide)
          iWindowState = (int)iWindowState | (int)Visio.VisWindowStates.visWSAnchorAutoHide;
        if (bVisible)
          iWindowState = (int)iWindowState | (int)Visio.VisWindowStates.visWSVisible;
        iWindowType = (int)Visio.VisWinTypes.visAnchorBarAddon;
        AddWindow(visActiveWindow, strCaption, iWindowState, iWindowType,
                  iLeft, iTop, iWidth, iHeight, strMergeID, strMergeClass, iMergePosition,
                  out visNewWindow);
        if (visNewWindow != null)
          {
          visWindowToAdd = visNewWindow;
          bSuccess = true;
          }
        }
      return bSuccess;
      }

    public static bool AddFormToAnchorWindow(Visio.Window anchorWindow, Form form)
      {

      int left, top, width, height;
      int windowHandle;
      bool bSuccess = false;

      try
        {

        // Show the form as a modeless dialog.
        form.Show();
        // Get the window handle of the form.
        windowHandle = form.Handle.ToInt32();
        // Set the form as a visible child window.
        if (NativeMethods.SetWindowLongW(windowHandle,
                                         NativeMethods.GWL_STYLE,
                                         NativeMethods.WS_CHILD | NativeMethods.WS_VISIBLE) != 0)
          {
          // La forme est maintenant une fenêtre enfant visible
          // Affectation de la anchor window en tant que parent de la forme
          if (NativeMethods.SetParent(windowHandle, anchorWindow.WindowHandle32) != 0)
            {
            // Le parent de la forme est maintenant la anchor window
            // Set the dock property of the form to fill, so that the form
            // automatically resizes to the size of the anchor bar.
            form.Dock = System.Windows.Forms.DockStyle.Fill;
            // Resize the anchor bar so it will refresh.
            anchorWindow.GetWindowRect(out left, out top, out width, out height);
            anchorWindow.SetWindowRect(left, top, width, height + 1);
            anchorWindow.SetWindowRect(left, top, width, height);
            bSuccess = true;
            }
          }
        }
      catch (Exception err)
        {
        }
      return bSuccess;
      }

    public static bool AddWindow(Visio.Window visWindow, string strCaption, object objState,
                                 object objType, int iLeft, int iTop, int iWidth, int iHeight,
                                 string strMergeID, string strMergeClass, int iMergePosition,
                                 out Visio.Window visCreatedWindow)
      {
      Visio.Windows visWindows;

      visCreatedWindow = null;
      try
        {
        visWindows = visWindow.Windows;
        visCreatedWindow = visWindows.Add(strCaption, objState, objType, null, null, null, null,
                                          strMergeID, strMergeClass, iMergePosition);
        }
      catch
        {
        return true;
        }
      return true;
      }

    internal static void GetColor(string strColor, out Color elementColor)
      {
      elementColor = Color.White;
      if (strColor.StartsWith("RGB"))
        {
        // c'est une couleur
        strColor = strColor.Replace("RGB(", "");
        strColor = strColor.Replace(")", "");
        string[] arTemp = strColor.Split(';');
        if (arTemp.Length == 3)
          {
          try
            {
            elementColor = Color.FromArgb(255, Convert.ToInt16(arTemp[0]), Convert.ToInt16(arTemp[1]),
                                          Convert.ToInt16(arTemp[2]));
            }
          catch
            {

            }

          }
        }
      }

    }
  }
