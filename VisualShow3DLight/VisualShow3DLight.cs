// VisualShow3DLight.cs
// Librairie VisualShow3DLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V0.6  |   ML		| 09/11/2017 12:00:00  |
//-------------------------------------------------------------------------//

using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;
using System.Configuration;

namespace VisualShow3DLight
  {

  /// <summary>
  /// Description résumée de VisualShow3DLight.
  /// </summary>
  public class VisualShow3DLight
    {
    // Configuration
    internal string strWebPath, strTemplatePath, strStencilPath;
    private Microsoft.Office.Interop.Visio.Application visApplication;
    float fScaleViewFactor = 1.0f;
    internal FrmBabylonScene frmBabylonPanel = null;
    internal FrmBabylonScene frmBabylonWindow = null;
    static internal Visio.Window visWindowBabylon = null;
    private Microsoft.Office.Interop.Visio.Document visDocument = null;
    private Microsoft.Office.Interop.Visio.Document visStencil = null;

    public VisualShow3DLight(Microsoft.Office.Interop.Visio.Application theApplication)
      {
      visApplication = theApplication;
      strWebPath = ConfigurationManager.AppSettings["WebPath"];
      strStencilPath = ConfigurationManager.AppSettings["StencilPath"];
      strTemplatePath = ConfigurationManager.AppSettings["TemplatePath"];
      }

    public void InitializeMember(Visio.Document visDocument)
      {
      if (this.visDocument != visDocument)
        {
        this.visDocument = visDocument;
        }
      }

    /// <summary>
    /// Création d'un nouveau document WVisioAddinBidon
    /// </summary>
    public void NewFile()
      {
      string strFullTemplateFilename, strFullStencilName;

      try
        {
        Cursor.Current = Cursors.WaitCursor;
        strFullTemplateFilename = Path.Combine(strTemplatePath, "VisualShow3DLight.vstx");
        strFullStencilName = Path.Combine(strStencilPath, "VisualShow3DLight.vssx");
        visDocument = visApplication.Documents.OpenEx(strFullTemplateFilename, (short)Visio.VisOpenSaveArgs.visOpenCopy);
        visStencil = visApplication.Documents.OpenEx(strFullStencilName,
          (short)Visio.VisOpenSaveArgs.visOpenRO
          + (short)Visio.VisOpenSaveArgs.visOpenMinimized
          + (short)Visio.VisOpenSaveArgs.visOpenDocked
          + (short)Visio.VisOpenSaveArgs.visOpenNoWorkspace);
        }
      catch (Exception excep)
        {
        }
      finally
        {
        Cursor.Current = Cursors.Default;
        }
      }

    public void OpenFile()
      {
      string strFullFilename;

      try
        {
        Cursor.Current = Cursors.WaitCursor;
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Title = "Open a diagram";
        openFileDialog.Filter = "Drawing(*.vsdx; *.vsd; *.vdx)| *.vsdx; *.vsd; *.vdx";
        openFileDialog.FilterIndex = 1;  // 1 based index
        if (openFileDialog.ShowDialog() == DialogResult.OK)
          {
          Cursor.Current = Cursors.WaitCursor;

          strFullFilename = openFileDialog.FileName;
          Cursor.Current = Cursors.WaitCursor;
          visDocument = visApplication.Documents.Open(strFullFilename);
          }
        }
      catch
        {
        }
      finally
        {
        Cursor.Current = Cursors.Default;
        //        bDocumentOpeningInProgress = false;
        }
      }

    public void SaveFile()
      {
      if (visDocument.Path == "")
        {
        // Not already saved
        SaveAsFile();
        }
      else
        {
        try
          {
          visDocument.Save();
          }
        catch
          {
          }
        }
      }

    public void SaveAsFile()
      {
      SaveFileDialog saveFileDialog = new SaveFileDialog();
      saveFileDialog.Title = "Save diagram";
      saveFileDialog.Filter = "Drawing(*.vsdx; *.vsd; *.vdx)| *.vsdx; *.vsd; *.vdx";
      saveFileDialog.FilterIndex = 1;  // 1 based index
      // Affiche la boite de sélection du logigramme
      if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
        string strFileName;

        strFileName = saveFileDialog.FileName;
        try
          {
          visDocument.SaveAs(strFileName);
          }
        catch
          {
          }
        finally
          {
          }
        }
      }


    public void CloseFile()
      {
      visDocument.Close();
      }
    public bool GetActivePageExtent(out float fMaxExtentX, out float fMaxExtentY, out float fScale, out bool bNoUnit)
      {
      Visio.Page visPage;
      double dblMaxExtentX, dblMaxExtentY, dblDrawingScale, dblScale;
      int iDrawingScaleType;
      double dblRulerOriginX, dblRulerOriginY;

      fMaxExtentX = 0.0f;
      fMaxExtentY = 0.0f;
      fScale = 1.0f;

      visPage = visApplication.ActivePage;
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "XRulerOrigin", (int)Visio.VisUnitCodes.visNumber,
                               out dblRulerOriginX);
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "YRulerOrigin", (int)Visio.VisUnitCodes.visNumber,
                               out dblRulerOriginY);
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "PageWidth", (int)Visio.VisUnitCodes.visNumber,
                               out dblMaxExtentX);
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "PageHeight", (int)Visio.VisUnitCodes.visNumber,
                               out dblMaxExtentY);
      fMaxExtentX = (float)(dblMaxExtentX);
      fMaxExtentY = (float)(dblMaxExtentY);
      VisualShow3DLightUtil.GetIntCellVal(visPage, "DrawingScaleType", out iDrawingScaleType);
      bNoUnit = (iDrawingScaleType == 0) ? true : false;
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "DrawingScale", (int)Visio.VisUnitCodes.visNumber,
                                 out dblDrawingScale);
      VisualShow3DLightUtil.GetDoubleCellVal(visPage, "PageScale", (int)Visio.VisUnitCodes.visNumber,
                                out dblScale);
      fScale = (float)(dblScale / dblDrawingScale) * fScaleViewFactor;
      return true;
      }

    public void ApplyShapeData()
      {
      Visio.Selection visSelection;
      int nbShapeSelected;
      ArrayList arOther = null;
      bool bOk = true;

      if (VisualShow3DLightUtil.GetActiveSelection(visApplication, out visSelection) == true)
        {
        nbShapeSelected = visSelection.Count;
        if (nbShapeSelected != 0)
          {
          arOther = new ArrayList();
          foreach (Visio.Shape visCurShape in visSelection)
            {
            arOther.Add(visCurShape);
            }
          }
        foreach (Visio.Shape visCurShape in arOther)
          {
          bOk &= VisualShow3DLightUtil.AddShapeDataRow(visCurShape, "VS3DMaterial", "Matériau",
                                          (int)Visio.VisCellVals.visPropTypeString, "", "Stone 01.bmp", 1036);
          bOk &= VisualShow3DLightUtil.AddShapeDataRow(visCurShape, "VS3DHeight", "Hauteur",
                                          (int)Visio.VisCellVals.visPropTypeNumber, "0.00 u", "1.00 cm", 1036);
          bOk &= VisualShow3DLightUtil.AddShapeDataRow(visCurShape, "VS3DElevation", "Elévation",
                                          (int)Visio.VisCellVals.visPropTypeNumber, "0.00 u", "0.00 cm", 1036);
          bOk &= VisualShow3DLightUtil.AddShapeDataRow(visCurShape, "VS3DGeometry", "Détails",
                                          (int)Visio.VisCellVals.visPropTypeBool, "", "FALSE", 1036);
          }
        }
      }

    public int GetNbElementsBabylonJS(bool bSelected)
      {
      Visio.Selection visSelection;
      int nbShapeSelected;
      Visio.Page visPage;
      int nCount = 0;

      if (bSelected)
        {
        //Récupération de la sélection active
        if (VisualShow3DLightUtil.GetActiveSelection(visApplication, out visSelection) == true)
          {
          nbShapeSelected = visSelection.Count;
          if (nbShapeSelected != 0)
            {
            foreach (Visio.Shape visCurShape in visSelection)
              {
              nCount++;
              }
            }
          }
        }
      else
        {
        visPage = visApplication.ActivePage;
        if (visPage != null)
          {
          foreach (Visio.Shape visCurShape in visPage.Shapes)
            {
            nCount++;
            }
          }
        }
      return nCount;
      }

    public bool GetElementsBabylonJS(out ArrayList ar3DObject, bool bSelected)
      {
      Visio.Selection visSelection;
      int nbShapeSelected;
      ArrayList arOther = null;
      Visio.Page visPage;

      ar3DObject = new ArrayList();
      arOther = new ArrayList();
      if (bSelected)
        {
        // Get Selection
        if (VisualShow3DLightUtil.GetActiveSelection(visApplication, out visSelection) == true)
          {
          nbShapeSelected = visSelection.Count;
          if (nbShapeSelected != 0)
            {
            foreach (Visio.Shape visCurShape in visSelection)
              {
              arOther.Add(visCurShape);
              }
            }
          }
        }
      else
        {
        visPage = visApplication.ActivePage;
        foreach (Visio.Shape visCurShape in visPage.Shapes)
          {
          arOther.Add(visCurShape);
          }
        }
      foreach (Visio.Shape visCurShape in arOther)
        {
        Visio3DObject cur3DObject = null;

        CreateVisio3DObjectFromBoundingBox(visCurShape, out cur3DObject);
        ar3DObject.Add(cur3DObject);
        }
      return (ar3DObject.Count != 0) ? true : false;
      }

    public void CreateVisio3DObjectFromBoundingBox(Visio.Shape visCurShape, out Visio3DObject visio3DObject)
      {
      PointF ptCurOrigin;
      double dblCurOrigX, dblCurOrigY, dblCurAngle, dblCurWidth, dblCurHeight, dblCurThickness, dblCurElevation;
      double dblCurLocPinX, dblCurLocPinY;
      float fCurAngle, fCurLength, fCurHeight, fCurThickness, fCurElevation;
      string strCurMaterial = "", strBoxTextureFileName = null, strFrontTextureFileName = null;
      Color elementColor = Color.White;

      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "PinX", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurOrigX);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "PinY", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurOrigY);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "LocPinX", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurLocPinX);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "LocPinY", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurLocPinY);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "Angle", (int)Visio.VisUnitCodes.visDegrees,
                                out dblCurAngle);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "Width", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurWidth);
      VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "Height", (int)Visio.VisUnitCodes.visNumber,
                                out dblCurThickness);
      try
        {
        VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "Prop.VS3DHeight", (int)Visio.VisUnitCodes.visNumber, out dblCurHeight);
        }
      catch
        {
        dblCurHeight = 0.0;
        }
      try
        {
        VisualShow3DLightUtil.GetDoubleCellVal(visCurShape, "Prop.VS3DElevation", (int)Visio.VisUnitCodes.visNumber, out dblCurElevation);
        }
      catch
        {
        dblCurElevation = 0.0;
        }
      try
        {
        VisualShow3DLightUtil.GetStringCellProp(visCurShape, "Prop.VS3DMaterial", out strBoxTextureFileName);
        }
      catch
        {
        strCurMaterial = "";
        }
      if (strCurMaterial == "")
        {
        VisualShow3DLightUtil.GetRGBCellVal(visCurShape, "FillForegnd", out strCurMaterial);
        VisualShow3DLightUtil.GetColor(strCurMaterial, out elementColor);
        }
      ptCurOrigin = new PointF((float)dblCurOrigX, (float)dblCurOrigY);
      fCurAngle = (float)dblCurAngle;
      fCurLength = (float)dblCurWidth;
      fCurHeight = (float)dblCurHeight;
      if (VisualShow3DLightUtil.IsOneD(visCurShape))
        fCurThickness = 0.0f;
      else
        fCurThickness = (float)dblCurThickness;
      fCurElevation = (float)dblCurElevation;
      visio3DObject = new Visio3DObject(ptCurOrigin, fCurAngle, fCurLength, fCurHeight, fCurThickness, fCurElevation,
                                      elementColor, Color.White, strBoxTextureFileName, strFrontTextureFileName);
      }


    internal void GenerateBabylonScene(bool bWall)
      {
      frmBabylonWindow = new FrmBabylonScene(visApplication);
      // Make it child of the Visio Window
      int windowHandle = frmBabylonWindow.Handle.ToInt32();
      VisualShow3DLightUtil.NativeMethods.SetParent(windowHandle, this.visApplication.WindowHandle32);
      frmBabylonWindow.Show();
      }

    internal void DisplayBabylon(string strTitle)
      {
      int left, top, width, height;

      this.visApplication.ActiveWindow.GetWindowRect(out left, out top, out width, out height);
      if (frmBabylonPanel == null)
        {
        if (VisualShow3DLightUtil.AddAnchorWindowToVisio(visApplication, strTitle,
                                          (int)Visio.VisWindowStates.visWSDockedBottom,
                                           false, true,
                                           0, 100, width, (int)(height * 0.5), "", "Vue 3D", 1,
                                           ref visWindowBabylon) == true)
          {
          frmBabylonPanel = new FrmBabylonScene(visApplication);
          if (VisualShow3DLightUtil.AddFormToAnchorWindow(visWindowBabylon, frmBabylonPanel) == true)
            {
            }
          else
            {
            frmBabylonPanel.Close();
            frmBabylonPanel = null;
            }
          }
        else
          {
          frmBabylonPanel.Dispose();
          frmBabylonPanel = null;
          }
        }
      else
        {
        }
      }

    internal void BabylonClose()
      {
      if (frmBabylonPanel != null)
        {
        frmBabylonPanel.Close();
        frmBabylonPanel = null;
        }
      }

    }
  }
