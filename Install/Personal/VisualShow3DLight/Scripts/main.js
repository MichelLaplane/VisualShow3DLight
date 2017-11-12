/// <reference path="babylon.js" />
var scene;
var camera;
document.addEventListener("DOMContentLoaded", startView, false);
function startView() {
    var canvas = document.getElementById("renderCanvas");
    var engine = new BABYLON.Engine(canvas, true);
    scene = new BABYLON.Scene(engine);
    var bAxisesDisplayed = false;
    if (window.external.IsAxises())
      ShowWorldAxises(scene, 5);
    camera = new BABYLON.ArcRotateCamera("Camera", -Math.PI / 2, Math.PI / 4, 20, BABYLON.Vector3.Zero(), scene);
    camera.attachControl(canvas, true);
    // Create a light located over the scene
    var light = new BABYLON.HemisphericLight("light1", new BABYLON.Vector3(0, 1, 0), scene);
    // Réglage de l'intensité
    light.intensity = 0.7;
    var box;
    //var sphere = BABYLON.Mesh.CreateSphere("sphere1", 16, 2, scene);
    // Déplacement vers le haut
    //sphere.position.y = 1;
    //var box;
    var elementExtentScale = { xExtent: 0.0, yExtent: 0.0, scaleFactor : 0.0 }
    window.external.GetExtentAndScale(elementExtentScale);
    var materialGround = new BABYLON.StandardMaterial("materialGround", scene);
    materialGround.alpha = 0.3;
    //    materialGround.wireframe = true;
    var ground = BABYLON.Mesh.CreateGround("ground1", elementExtentScale.xExtent, elementExtentScale.yExtent, 2, scene);
    ground.material = materialGround;
    DrawVisioDiagram(scene);
    engine.runRenderLoop(function ()
      {
      if (window.external.IsDynamic())
      {
        DynamicDrawVisioDiagram(scene);
        if (window.external.IsAxises())
        {
          if (!bAxisesDisplayed)
          {
            ShowWorldAxises(scene, 5);
            bAxisesDisplayed = true;
          }
        }
        else
        {
          if (bAxisesDisplayed)
          {
            HideWorldAxises(scene);
            bAxisesDisplayed = false;
          }
        }
          light.setEnabled(window.external.IsLight());
        scene.render();
        }
      else
        {
        scene.render();
        }
    });

    function DrawVisioDiagram(scene)
    {
      var nbElement = window.external.GetNbElements();
      for (var i = 0; i < nbElement; i++)
      {
        var elementShape = { x: 0, y: 0, height: 0, width: 0, depth: 0, elevation: 0, angle: 0, color: 0, texture : 0 };
        window.external.GetElements(elementShape, i);
        var box = BABYLON.MeshBuilder.CreateBox("box" + i, {
          height: elementShape.height,
          width: elementShape.width,
          depth: elementShape.depth
        }, scene);
        var arColor = elementShape.color.split(",");
        var materialBox = new BABYLON.StandardMaterial("color", scene);
        materialBox.diffuseColor = new BABYLON.Color3(arColor[0] / 255,
                                                        arColor[1] / 255,
                                                        arColor[2] / 255);
        if (elementShape.texture != "")
            materialBox.diffuseTexture = new BABYLON.Texture("../Textures/" + elementShape.texture, scene);
        box.material = materialBox;
        box.position.x = elementShape.x;
        box.position.z = elementShape.y;
        box.position.y = elementShape.elevation;
        box.rotation.y = elementShape.angle * Math.PI / 180;
        //alert(box.position.x.toFixed(2) + "," + box.position.z.toFixed(2) + "," + box.position.y.toFixed(2));
        box.translate(new BABYLON.Vector3(-1, 0, 0), elementExtentScale.xExtent * 0.5, BABYLON.Space.WORLD);
        box.translate(new BABYLON.Vector3(0, 0, -1), elementExtentScale.yExtent * 0.5, BABYLON.Space.WORLD);
        if (elementShape.height == 0)
          {
            // height is not available in Shape probably because the user had not applied ShapeData
            // so we'll use the default height provided by MeshBuilder
            box.translate(new BABYLON.Vector3(0, 1, 0), 0.5, BABYLON.Space.WORLD);
          }
        else
          box.translate(new BABYLON.Vector3(0, 1, 0), elementShape.height * 0.5, BABYLON.Space.WORLD);
      }

    }

    function DynamicDrawVisioDiagram(scene)
    {
        var nbElement = window.external.GetNbElements();
        for (var i = 0; i < nbElement; i++)
          {
            var curBox = scene.getMeshByName("box" + i);
            if (curBox)
              {
              curBox.dispose();
              }
            var elementShape = { x: 0, y: 0, height: 0, width: 0, depth: 0, elevation: 0, angle: 0, color: 0, texture: 0  };
            window.external.GetElements(elementShape, i);
            var box = BABYLON.MeshBuilder.CreateBox("box" + i, {
              height: elementShape.height,
              width: elementShape.width,
              depth: elementShape.depth
            }, scene);
            var arColor = elementShape.color.split(",");
            var materialBox = new BABYLON.StandardMaterial("color", scene);
            materialBox.diffuseColor = new BABYLON.Color3(arColor[0] / 255,
                                                            arColor[1] / 255,
                                                            arColor[2] / 255);
            if (elementShape.texture != "")
              materialBox.diffuseTexture = new BABYLON.Texture("../Textures/" + elementShape.texture, scene);
            box.material = materialBox;
            box.position.x = elementShape.x;
            box.position.z = elementShape.y;
            box.position.y = elementShape.elevation;
            box.rotation.y = elementShape.angle * Math.PI / 180;
            //alert(box.position.x.toFixed(2) + "," + box.position.z.toFixed(2) + "," + box.position.y.toFixed(2));
            box.translate(new BABYLON.Vector3(-1, 0, 0), elementExtentScale.xExtent * 0.5, BABYLON.Space.WORLD);
            box.translate(new BABYLON.Vector3(0, 0, -1), elementExtentScale.yExtent * 0.5, BABYLON.Space.WORLD);
            if (elementShape.height == 0)
            {
              // height is not available in Shape probably because the user had not applied ShapeData
              // so we'll use the default height provided by MeshBuilder
              box.translate(new BABYLON.Vector3(0, 1, 0), 0.5, BABYLON.Space.WORLD);
            }
            else
              box.translate(new BABYLON.Vector3(0, 1, 0), elementShape.height * 0.5, BABYLON.Space.WORLD);
        }


        function GenerateAll()
        {
          alert("I'm in GenerateAll function")
          DrawVisioDiagram(scene);
        }



    }
}