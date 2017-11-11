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
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VisualShow3DLight
  {
  [ComVisible(true)]
  public partial class FrmBabylonScene : Form
    {
    bool bDraw;

    public FrmBabylonScene(Microsoft.Office.Interop.Visio.Application visApp)
      {

      InitializeComponent();
      chkLight.Checked = true;
      webBrowserBabylon.ObjectForScripting  = this;
      }

    private void btnFloorPlanAll_Click(object sender, EventArgs e)
      {
      // This does not work because the function is not in the HTML File
      // Don't find a solution for that
      webBrowserBabylon.Document.InvokeScript("GenerateAll");
      }


    public int GetNbElements()
      {
      return ThisAddIn.addinApplication.GetNbElementsBabylonJS(false);
      }

    public bool IsDynamic()
      {
      return chkDynamic.Checked;
      }

    public bool IsAxises()
      {
      return chkAxises.Checked;
      }

    public bool IsLight()
      {
      return chkLight.Checked;
      }

    public void GetExtentAndScale(dynamic element)
      {
      float fMaxExtentX, fMaxExtentY, fScale;
      bool bNoUnit;

      ThisAddIn.addinApplication.GetActivePageExtent(out fMaxExtentX, out fMaxExtentY, out fScale, out bNoUnit);
      element.xExtent = fMaxExtentX;
      element.yExtent = fMaxExtentY;
      element.scaleFactor = fScale;
      }

    public void GetElements(dynamic element, int i)
      {
      ArrayList ar3DObject;

      if (ThisAddIn.addinApplication.GetElementsBabylonJS(out ar3DObject, false))
        {
        if(i < ar3DObject.Count)
          {
          element.x = ((Visio3DObject)ar3DObject[i]).ptOrigin.X;
          element.y = ((Visio3DObject)ar3DObject[i]).ptOrigin.Y;
          element.height = ((Visio3DObject)ar3DObject[i]).fHeight;
          element.width = ((Visio3DObject)ar3DObject[i]).fLength;
          element.depth = ((Visio3DObject)ar3DObject[i]).fThickness;
          element.elevation = ((Visio3DObject)ar3DObject[i]).fElevation;
          element.angle = -((Visio3DObject)ar3DObject[i]).fAngle;
          element.color = ((Visio3DObject)ar3DObject[i]).boxColor.R.ToString() + "," +
                          ((Visio3DObject)ar3DObject[i]).boxColor.G.ToString() + "," +
                          ((Visio3DObject)ar3DObject[i]).boxColor.B.ToString();
          element.texture = ((Visio3DObject)ar3DObject[i]).strBoxTexture;
          }
        }
      }

    public bool getObject()
      {
      return bDraw;
      }

    public bool AltitudeUp()
      {
      webBrowserBabylon.Document.InvokeScript("AltitudeUp");
      return true;
      }

    public bool AltitudeDown()
      {
      webBrowserBabylon.Document.InvokeScript("AltitudeDown");
      return true;
      }

    public bool AzimuthInc()
      {
      webBrowserBabylon.Document.InvokeScript("AzimuthInc");
      return true;
      }

    public bool AzimuthDec()
      {
      webBrowserBabylon.Document.InvokeScript("AzimuthDec");
      return true;
      }

    public bool PanLeft()
      {
      webBrowserBabylon.Document.InvokeScript("PanLeft");
      return true;
      }

    public bool PanRight()
      {
      webBrowserBabylon.Document.InvokeScript("PanRight");
      return true;
      }

    public bool ZoomIn()
      {
      webBrowserBabylon.Document.InvokeScript("ZoomIn");
      return true;
      }

    public bool ZoomOut()
      {
      webBrowserBabylon.Document.InvokeScript("ZoomOut");
      return true;
      }
    
    //private void webBrowserBabylon_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
    //  {
    //  //when the form is load cursor focus on the Web browser control.
    //  webBrowserBabylon.Focus();
    //  }

    private void FrmBabylonScene_Load(object sender, EventArgs e)
      {
      webBrowserBabylon.Navigate(Path.Combine(ThisAddIn.addinApplication.strWebPath, "index.html"));
      }
    }


  }
