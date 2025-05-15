/**
 * Demo data generator for testing Excel output
 * This provides simple sample data to test the Make-Ready Report fixes
 * without needing real SPIDA and Katapult data files
 */

/**
 * Generate a simple SPIDA data structure for testing
 */
export function generateDemoSpidaData(): any {
  return {
    "leads": [
      {
        "locations": [
          createDemoPoleLocation("PL410620"),
          createDemoPoleLocation("PL398491")
        ]
      }
    ],
    "clientData": {
      "poles": [
        {
          "aliases": [{
            "id": "demo-pole-id-1"
          }],
          "species": "Southern Pine",
          "classOfPole": "40-4"
        }
      ]
    }
  };
}

/**
 * Create a demo pole location for SPIDA data
 */
function createDemoPoleLocation(poleId: string): any {
  return {
    "label": poleId,
    "designs": [
      {
        "name": "Measured Design",
        "structure": {
          "pole": {
            "clientItemAlias": "40-4",
            "clientItem": {
              "id": "demo-pole-id-1",
              "species": "Southern Pine"
            }
          },
          "wires": [
            {
              "id": "wire-1",
              "owner": { "id": "CPS" },
              "clientItem": { "description": "Neutral", "type": "PRIMARY" },
              "attachmentHeight": { "value": 9.91 } // ~32'6"
            },
            {
              "id": "wire-2",
              "owner": { "id": "CPS" },
              "clientItem": { "description": "Primary", "type": "PRIMARY" },
              "attachmentHeight": { "value": 10.7 } // ~35'0"
            },
            {
              "id": "wire-3",
              "owner": { "id": "AT&T" },
              "clientItem": { "description": "Fiber Optic Com", "type": "COMMUNICATION" },
              "attachmentHeight": { "value": 7.7 } // ~25'3"
            }
          ],
          "equipments": [],
          "wireEndPoints": [
            {
              "type": "NEXT_POLE",
              "structureLabel": poleId === "PL410620" ? "PL398491" : "PL410620",
              "wires": ["wire-1", "wire-2", "wire-3"],
              "direction": 45
            }
          ]
        }
      },
      {
        "name": "Recommended Design",
        "structure": {
          "pole": {
            "clientItemAlias": "40-4",
            "clientItem": {
              "id": "demo-pole-id-1",
              "species": "Southern Pine"
            }
          },
          "wires": [
            {
              "id": "wire-1",
              "owner": { "id": "CPS" },
              "clientItem": { "description": "Neutral", "type": "PRIMARY" },
              "attachmentHeight": { "value": 10.05 } // ~33'0"
            },
            {
              "id": "wire-2",
              "owner": { "id": "CPS" },
              "clientItem": { "description": "Primary", "type": "PRIMARY" },
              "attachmentHeight": { "value": 10.7 } // ~35'0" (same as existing)
            },
            {
              "id": "wire-3",
              "owner": { "id": "AT&T" },
              "clientItem": { "description": "Fiber Optic Com", "type": "COMMUNICATION" },
              "attachmentHeight": { "value": 7.7 } // ~25'3" (same as existing)
            }
          ],
          "equipments": [],
          "wireEndPoints": [
            {
              "type": "NEXT_POLE",
              "structureLabel": poleId === "PL410620" ? "PL398491" : "PL410620",
              "wires": ["wire-1", "wire-2", "wire-3"],
              "direction": 45
            }
          ]
        },
        "analysis": [
          {
            "analysisCaseDetails": {
              "name": "Light - Grade C",
              "constructionGrade": "C"
            },
            "results": [
              {
                "component": "Pole",
                "analysisType": "STRESS",
                "actual": 85.2,
                "unit": "PERCENT"
              }
            ]
          }
        ]
      }
    ]
  };
}

/**
 * Generate a simple Katapult data structure for testing
 */
export function generateDemoKatapultData(): any {
  return {
    "nodes": [
      {
        "id": "node-1",
        "properties": {
          "PoleNumber": "PL410620"
        }
      },
      {
        "id": "node-2",
        "properties": {
          "PoleNumber": "PL398491"
        }
      }
    ],
    "connections": {
      "conn-1": {
        "node_id_1": "node-1",
        "node_id_2": "node-2",
        "button_type": "aerial_path",
        "sections": {
          "section-1": {
            "photos": {
              "photo-1": true
            }
          }
        }
      }
    },
    "photo_summary": {
      "photo-1": {
        "photofirst_data": {
          "wire": {
            "instance-1": {
              "_measured_height": 270, // 22'-6"
              "trace_id": "trace-1"
            },
            "instance-2": {
              "_measured_height": 342, // 28'-6"
              "trace_id": "trace-2"
            },
            "instance-3": {
              "_measured_height": 254, // 21'-2"
              "trace_id": "trace-3"
            }
          }
        }
      }
    },
    "traces": {
      "trace_data": {
        "trace-1": {
          "company": "AT&T",
          "cable_type": "Fiber Optic Com"
        },
        "trace-2": {
          "company": "CPS",
          "cable_type": "Primary"
        },
        "trace-3": {
          "company": "CPS",
          "cable_type": "Supply Fiber"
        }
      }
    }
  };
}
