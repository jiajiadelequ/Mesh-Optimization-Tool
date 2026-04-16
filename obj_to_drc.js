const fs = require('fs');
const path = require('path');
const draco3d = require('draco3d');

function parseObj(text) {
  const vertices = [];
  const faces = [];
  const lines = text.split(/\r?\n/);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) {
      continue;
    }

    if (line.startsWith('v ')) {
      const parts = line.split(/\s+/);
      if (parts.length >= 4) {
        vertices.push(Number(parts[1]), Number(parts[2]), Number(parts[3]));
      }
      continue;
    }

    if (line.startsWith('f ')) {
      const parts = line.split(/\s+/).slice(1);
      if (parts.length < 3) {
        continue;
      }
      const indices = parts.map((token) => {
        const idx = token.split('/')[0];
        return Number(idx) - 1;
      });
      for (let i = 1; i < indices.length - 1; i += 1) {
        faces.push(indices[0], indices[i], indices[i + 1]);
      }
    }
  }

  if (vertices.length === 0 || faces.length === 0) {
    throw new Error('OBJ has no vertices or faces');
  }

  return {
    vertices: new Float32Array(vertices),
    faces: new Uint32Array(faces),
  };
}

async function main() {
  const [inputPath, outputPath] = process.argv.slice(2);
  if (!inputPath || !outputPath) {
    throw new Error('Usage: node obj_to_drc.js <input.obj> <output.drc>');
  }

  const objText = fs.readFileSync(inputPath, 'utf8');
  const parsed = parseObj(objText);
  const encoderModule = await draco3d.createEncoderModule({});

  const meshBuilder = new encoderModule.MeshBuilder();
  const mesh = new encoderModule.Mesh();
  const encoder = new encoderModule.Encoder();

  try {
    const numFaces = parsed.faces.length / 3;
    const numPoints = parsed.vertices.length / 3;

    meshBuilder.AddFacesToMesh(mesh, numFaces, parsed.faces);
    meshBuilder.AddFloatAttributeToMesh(
      mesh,
      encoderModule.POSITION,
      numPoints,
      3,
      parsed.vertices
    );

    encoder.SetSpeedOptions(5, 5);
    encoder.SetAttributeQuantization(encoderModule.POSITION, 11);
    encoder.SetEncodingMethod(encoderModule.MESH_EDGEBREAKER_ENCODING);

    const encodedData = new encoderModule.DracoInt8Array();
    const encodedLen = encoder.EncodeMeshToDracoBuffer(mesh, encodedData);
    if (encodedLen <= 0) {
      throw new Error('Draco encoding failed');
    }

    const outputBuffer = Buffer.alloc(encodedLen);
    for (let i = 0; i < encodedLen; i += 1) {
      outputBuffer[i] = encodedData.GetValue(i);
    }

    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, outputBuffer);

    encoderModule.destroy(encodedData);
  } finally {
    encoderModule.destroy(encoder);
    encoderModule.destroy(mesh);
    encoderModule.destroy(meshBuilder);
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
