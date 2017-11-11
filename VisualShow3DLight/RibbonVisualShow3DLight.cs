// RibbonVisualShow3DLight.cs
// Librairie VisualShow3DLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V1.0  |   ML		| 00/00/2011 15:52:49  |
//-------------------------------------------------------------------------//

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;


namespace VisualShow3DLight
{
  [ComVisible(true)]
  public class RibbonVisualShow3DLight : Office.IRibbonExtensibility
  {
    internal Office.IRibbonUI ribbon;

    public RibbonVisualShow3DLight()
    {
    }

    #region Membres IRibbonExtensibility

    public string GetCustomUI(string ribbonID)
    {
      return GetResourceText("VisualShow3DLight.RibbonVisualShow3DLight.xml");
    }

    #endregion

    #region Rappels du ruban
    //Créez des méthodes de rappel ici. Pour plus d'informations sur l'ajout de méthodes de rappel, sélectionnez l'élément XML Ruban dans l'Explorateur de solutions, puis appuyez sur F1

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
      this.ribbon = ribbonUI;
    }

    #endregion

    #region Programmes d'assistance

    private static string GetResourceText(string resourceName)
    {
      Assembly asm = Assembly.GetExecutingAssembly();
      string[] resourceNames = asm.GetManifestResourceNames();
      for (int i = 0; i < resourceNames.Length; ++i)
      {
        if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
        {
          using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
          {
            if (resourceReader != null)
            {
              return resourceReader.ReadToEnd();
            }
          }
        }
      }
      return null;
    }

    #endregion

    public void OnAction(Office.IRibbonControl control)
    {
      switch (control.Id)
      {
        case "btnApplyShapeData":
          // Application des données de forme
          ThisAddIn.addinApplication.ApplyShapeData();
          break;
        case "btnSceneGenerateWithBabylonJS":
          // BabylonJS in Popup Window
          ThisAddIn.addinApplication.GenerateBabylonScene(true);
          break;
        case "btnWindowSceneGenerateWithBabylonJS":
          // BabylonJS in panel Window
          ThisAddIn.addinApplication.DisplayBabylon("Vue 3D");
          break;
        //Backstage
        case "btnProjectNew":
          ThisAddIn.addinApplication.NewFile();
          break;
        case "btnProjectOpen":
          ThisAddIn.addinApplication.OpenFile();
          break;
        case "btnProjectSave":
          ThisAddIn.addinApplication.SaveFile();
          break;
        case "btnProjectSaveAs":
          ThisAddIn.addinApplication.SaveAsFile();
          break;
        case "btnProjectClose":
          ThisAddIn.addinApplication.CloseFile();
          break;
        default:
          break;
        }
      }

    public bool GetEnabled(Microsoft.Office.Core.IRibbonControl control)
    {
      bool bRetour = true;

      switch (control.Id)
      {
        case "btnNew":
          {
            bRetour = true;
            break;
          }
        default:
          break;
      }
      return bRetour;
    }

    /// <summary>
    /// Renvoi une image à l'appelant pour un bouton.
    /// </summary>
    /// <param name="control"></param>
    /// <returns></returns>
    public System.Drawing.Bitmap GetImage(Microsoft.Office.Core.IRibbonControl control)
    {
      switch (control.Id)
      {
        // Backstage
        case "taskCatFiles":
          return Properties.Resources.FileManagement64;
        case "btnProjectNew":
          return Properties.Resources.ProjectNew64;
        case "btnProjectOpen":
          return Properties.Resources.ProjectOpen64;
        case "btnProjectSave":
          return Properties.Resources.ProjectSave64;
        case "btnProjectSaveAs":
          return Properties.Resources.ProjectSaveAs64;
        case "btnProjectClose":
          return Properties.Resources.ProjectClose64;
        // Ruban
        case "btnApplyShapeData":
          // Application des données de forme
          return Properties.Resources.ApplyShapeData;
        case "btnSceneGenerateWithBabylonJS":
          // BabylonJS in Popup Window
          return Properties.Resources.Vue3D32;
        case "btnWindowSceneGenerateWithBabylonJS":
          // BabylonJS in panel Window
          return Properties.Resources.Panel3D32;
        default:
          break;

          //case "buttonGetHelp":
          //      {
          //      return Properties.Resources.SelfSupportPH;
          //      }
          //case "buttonTemplate1":
          //case "imageControl1":
          //        {
          //        return Properties.Resources.TemplateIcon1;
          //        }
      }

      // we should not get here for these buttons
      return null;
    }

    public string GetLabel(Microsoft.Office.Core.IRibbonControl control)
      {
      Assembly applicationAssembly;
      string strVersion = "", strVersions = "";

      switch (control.Id)
        {
        case "labelBuildInfo":
          applicationAssembly = Assembly.GetCallingAssembly();
          strVersion = "Version : " + applicationAssembly.GetName().Version.ToString();
          break;
        default:
          break;
        }
      return strVersion;
      }

    }
}
