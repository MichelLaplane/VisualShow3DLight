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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;
#if VISIOASM
using VisioAsm;
using VisMeth = VisioAsm.VLMethods;
using VisCst = VisioAsm.VLConstants;
#endif
#if STRINGASM
using StringAsm;
#endif

namespace VisualShow3DLight
{
  public class Visio3DObject
    {
    public PointF ptOrigin;
    public float fAngle;
    public float fLength;
    public float fHeight;
    public float fThickness;
    public float fElevation;
    public Color frontColor;
    public Color boxColor;
    public string strFrontTexture;
    public string strBoxTexture;

    public Visio3DObject(PointF ptParamOrigin, float fParamAngle, float fParamLength, float fParamHeight,
                         float fParamThickness, float fParamElevation, Color paramBoxColor, Color paramFrontColor,
                         string strParamBoxTexture, string strParamFrontTexture)
      {
      ptOrigin = ptParamOrigin;
      fAngle = fParamAngle;
      fLength = fParamLength;
      fHeight = fParamHeight;
      fThickness = fParamThickness;
      fElevation = fParamElevation;
      boxColor = paramBoxColor;
      frontColor = paramFrontColor;
      strBoxTexture = strParamBoxTexture;
      strFrontTexture = strParamFrontTexture;
      }

    }
  }
