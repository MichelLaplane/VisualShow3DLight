// ThisAddin.cs
// Librairie VisualShow3DLight
// Copyright © ShareVisual Michel LAPLANE
// All rights reserved.

//-------------------------------------------------------------------------//
//					TABLEAU DE BORD DES MISES A JOUR
//-------------------------------------------------------------------------//
//Modifié: V0.6  |   ML		| 09/11/2017 12:00:00  |
//-------------------------------------------------------------------------//
using System;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisualShow3DLight
  {

  public partial class ThisAddIn
  {
    private Visio.Application visApplication;
    static public VisualShow3DLight addinApplication;
    static internal RibbonVisualShow3DLight ribbonApplication;
    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {

      if (visApplication == null)
      {
        visApplication = (Microsoft.Office.Interop.Visio.Application)this.Application;
      }

      addinApplication = new VisualShow3DLight(visApplication);
      // Event subscription
      Globals.ThisAddIn.Application.BeforeWindowClosed += new Microsoft.Office.Interop.Visio.EApplication_BeforeWindowClosedEventHandler(OnBeforeWindowClosed);
      Globals.ThisAddIn.Application.OnKeystrokeMessageForAddon += new Microsoft.Office.Interop.Visio.EApplication_OnKeystrokeMessageForAddonEventHandler(OnKeystrokeMessageForAddon);

      }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
      try
      {
        visApplication = null;
      }
      catch
      {
      }
    }

    #region vsto ribbon support

    /// <summary>
    /// Retrieve Ribbon when Visio loads the VSTO
    /// </summary>
    /// <returns></returns>
    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
    {
      ThisAddIn.ribbonApplication = new RibbonVisualShow3DLight();
      return ThisAddIn.ribbonApplication;
    }


    #endregion

    #region Code généré par VSTO

    /// <summary>
    /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
    /// le contenu de cette méthode avec l'éditeur de code.
    /// </summary>
    private void InternalStartup()
    {
      this.Startup += new System.EventHandler(ThisAddIn_Startup);
      this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    }

    #endregion

    public void OnBeforeWindowClosed(Visio.Window visWindow)
      {
      string strCaption;

      strCaption = visWindow.Caption;
      if (strCaption.Contains("Vue 3D"))
        {
        addinApplication.BabylonClose();
        }
      }

    public bool OnKeystrokeMessageForAddon(Visio.MSGWrap msgWrap)
      {
      System.Windows.Forms.Message msg = new System.Windows.Forms.Message();
      msg.Msg = msgWrap.message;
      msg.WParam = (IntPtr)msgWrap.wParam;
      msg.LParam = (IntPtr)msgWrap.lParam;
      ThisAddIn.addinApplication.frmBabylonPanel.webBrowserBabylon.PreProcessMessage(ref msg);
      return true;
      }

    }
  }
