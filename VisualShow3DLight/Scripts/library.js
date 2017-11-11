// Utilities Part of BabylonJS samples

function ShowWorldAxises(scene,size)
{
  var makeTextPlane = function (id, text, color, size)
  {
    var dynamicTexture = new BABYLON.DynamicTexture("DynamicTexture"+id, 50, scene, true);
    dynamicTexture.hasAlpha = true;
    dynamicTexture.drawText(text, 5, 40, "bold 36px Arial", color, "transparent", true);
    var plane = BABYLON.Mesh.CreatePlane("TextPlane" + id, size, scene, true);
    plane.material = new BABYLON.StandardMaterial("TextPlaneMaterial"+id, scene);
    plane.material.backFaceCulling = false;
    plane.material.specularColor = new BABYLON.Color3(0, 0, 0);
    plane.material.diffuseTexture = dynamicTexture;
    return plane;
  };
  var axisX = BABYLON.Mesh.CreateLines("axisX", [
    BABYLON.Vector3.Zero(), new BABYLON.Vector3(size, 0, 0), new BABYLON.Vector3(size * 0.95, 0.05 * size, 0),
    new BABYLON.Vector3(size, 0, 0), new BABYLON.Vector3(size * 0.95, -0.05 * size, 0)
  ], scene);
  axisX.color = new BABYLON.Color3(1, 0, 0);
  var xChar = makeTextPlane("1","X", "red", size / 10);
  xChar.position = new BABYLON.Vector3(0.9 * size, -0.05 * size, 0);
  var axisY = BABYLON.Mesh.CreateLines("axisY", [
    BABYLON.Vector3.Zero(), new BABYLON.Vector3(0, size, 0), new BABYLON.Vector3(-0.05 * size, size * 0.95, 0),
    new BABYLON.Vector3(0, size, 0), new BABYLON.Vector3(0.05 * size, size * 0.95, 0)
  ], scene);
  axisY.color = new BABYLON.Color3(0, 1, 0);
  var yChar = makeTextPlane("2","Y", "green", size / 10);
  yChar.position = new BABYLON.Vector3(0, 0.9 * size, -0.05 * size);
  var axisZ = BABYLON.Mesh.CreateLines("axisZ", [
    BABYLON.Vector3.Zero(), new BABYLON.Vector3(0, 0, size), new BABYLON.Vector3(0, -0.05 * size, size * 0.95),
    new BABYLON.Vector3(0, 0, size), new BABYLON.Vector3(0, 0.05 * size, size * 0.95)
  ], scene);
  axisZ.color = new BABYLON.Color3(0, 0, 1);
  var zChar = makeTextPlane("3","Z", "blue", size / 10);
  zChar.position = new BABYLON.Vector3(0, 0.05 * size, 0.9 * size);
}

function HideWorldAxises(scene)
{
  var curElement = scene.getMeshByName("axisX");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("axisY");
  if (curElement)
  {
    curElement.dispose();
  }
   curElement = scene.getMeshByName("axisZ");
  if (curElement)
  {
    curElement.dispose();
   }
  curElement = scene.getMeshByName("TextPlane1");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("TextPlane2");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("TextPlane3");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("TextPlaneMaterial1");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("TextPlaneMaterial2");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("TextPlaneMaterial3");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("DynamicTexture1");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("DynamicTexture2");
  if (curElement)
  {
    curElement.dispose();
  }
  curElement = scene.getMeshByName("DynamicTexture3");
  if (curElement)
  {
    curElement.dispose();
  }
}