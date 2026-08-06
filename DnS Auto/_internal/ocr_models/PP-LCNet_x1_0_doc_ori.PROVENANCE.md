# PP-LCNet_x1_0_doc_ori model provenance

- Model: PP-LCNet_x1_0_doc_ori
- Upstream: PaddleX / PaddleOCR (Baidu)
- License: Apache-2.0
- Official Paddle inference archive: https://paddle-model-ecology.bj.bcebos.com/paddlex/official_inference_model/paddle3.0.0/PP-LCNet_x1_0_doc_ori_infer.tar
- Official archive SHA-256: 282337df5c41f7cdf8dacd5acf71fddfdc10218399f4b318463c17f4eae96c97
- Product ONNX SHA-256: 1db9914a3beb04181fde445b2fef96b850072f89a2fa8aa71ebef4ed03b8074f
- Conversion: official PaddleX `paddlex --paddle2onnx` command on Ubuntu GitHub Actions
- Validation: 16 synthetic document-orientation inputs; all Paddle/ONNX classes equal; maximum absolute output difference 3.5762786865234375e-07

The ONNX file was accepted only after direct numerical comparison with the official
Paddle inference model. A third-party candidate model was rejected and is not shipped.