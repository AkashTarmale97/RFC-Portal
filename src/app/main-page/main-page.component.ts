import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MatButtonModule } from '@angular/material/button';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatLabel } from '@angular/material/form-field';
import { MatFormField } from '@angular/material/form-field';
import { MatSelectModule } from '@angular/material/select';
import {
  Document, Packer, Paragraph, TextRun, Header, Footer, ImageRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  VerticalAlign,
  HeightRule,
  PageNumber,
  TableRowHeight
} from 'docx';
import { saveAs } from 'file-saver';
import { Router } from '@angular/router';



@Component({
  selector: 'app-main-page',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    ReactiveFormsModule,
    MatButtonModule,
    MatFormFieldModule,
    MatLabel,
    MatFormField,
    MatSelectModule
  ],
  templateUrl: './main-page.component.html',
  styleUrl: './main-page.component.css'
})
export class MainPageComponent {
  activeTab: string = "users";

  formData = {
    application_name: '',
    description: '',
    email: '',
    details: '',
    initiation_date: '',
    contact_number: '',
    change_execution: '',
    nature_of_change: '',
    implementation_date: '',
    type_apicode: false,
    type_architectural: false,
    type_infra: false,
    type_audit_compliance:false,
    type_hardware:false,
    type_design:false,
    type_configuration:false,
    type_quality_changes:false,
    type_oem_recommendation:false,
    type_network:false,
    type_security:false,
    type_Software:false,
    type_upgradation:false,
    type_others:false,
    change_category: '',
    initiation_date_label: '',
    ObjectiveofChange:''

  };


// // TO hide past dates
mintoDate: string | undefined;

ngOnInit(): void {
  const today = new Date();
  this.mintoDate = today.toISOString().split('T')[0]; // format: yyyy-MM-dd
  this.formData.initiation_date = this.mintoDate;
}


  constructor(private router: Router) { }

  companyLogoBase64 = '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';

  onSubmit() {
    alert("RFC generated Successfully!!!!!");
    const logoImage = new ImageRun({
      data: Uint8Array.from(atob(this.companyLogoBase64), c => c.charCodeAt(0)),
      transformation: {
        width: 150,
        height: 40,
      },
      type: "png", // or "jpeg" depending on your image
    });

    const header = new Header({
      children: [
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
            bottom: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
            left: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
            right: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
            insideHorizontal: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
            insideVertical: { style: BorderStyle.DOUBLE, size: 3, color: "000000" },
          },
          rows: [
            new TableRow({
              height: {
                value: 600, // controls row height
                rule: 'exact',
                // rule: HeightRule.ATLEAST,
              },
              children: [
                new TableCell({
                  children: [new Paragraph({ children: [logoImage] })],
                  // margins: {
                  //   top: 100,      // 100 twips ≈ 0.07 inches
                  //   bottom: 100,
                  //   left: 100,
                  //   right: 100,
                  // },
                  width: { size: 40, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: 'Change Request Form', bold: true, size: 24 })],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  // margins: {
                  //   top: 100,      // 100 twips ≈ 0.07 inches
                  //   bottom: 100,
                  //   left: 100,
                  //   right: 100,
                  // },
                  width: { size: 34, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: '(EIS)', bold: true, size: 24 })],
                      alignment: AlignmentType.CENTER,
                    }),
                  ],
                  // margins: {
                  //   top: 100,      // 100 twips ≈ 0.07 inches
                  //   bottom: 100,
                  //   left: 100,
                  //   right: 100,
                  // },
                  width: { size: 33, type: WidthType.PERCENTAGE },

                  verticalAlign: VerticalAlign.CENTER,
                }),
              ],
            }),
          ],
        }),
      ],
    });

    const footer = new Footer({
      children: [
        new Table({
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
          borders: {
            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
          },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  width: { size: 33, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      children: [
                        new TextRun({
                          text: "SBI Confidential",
                          size: 20,
                          italics: true,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 34, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun("Page "),
                        new TextRun({
                          children: [PageNumber.CURRENT],
                        }),
                        new TextRun(" of "),
                        new TextRun({
                          children: [PageNumber.TOTAL_PAGES],
                        }),
                      ],
                    })
                  ],
                }),
                new TableCell({
                  width: { size: 33, type: WidthType.PERCENTAGE },
                  verticalAlign: VerticalAlign.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      children: [
                        new TextRun({
                          text: "Version 1.0",
                          size: 20,
                          italics: true,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });


    // Generate CRF Number
    const today = new Date();
    const day = today.getDate();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();
    let sequentialId = parseInt(localStorage.getItem("crfCount") || "0", 10) + 1;
    localStorage.setItem("crfCount", sequentialId.toString()); // Can be auto-incremented if needed

    const crfNumber = `CRQ${day}${month}${year}${sequentialId}`;

    // Create title row with CRF box
    const titleAndCRFRow = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'PART 1: CHANGE REQUEST (To be completed by Change Requester)',
                      bold: true,
                      size: 20,
                    }),
                  ],
                }),
              ],
              width: { size: 70, type: WidthType.PERCENTAGE },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: `CRF No - ${crfNumber}`, bold: true }),
                  ],
                }),
              ],
              width: { size: 30, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
                bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
                left: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
                right: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
              },
            }),
          ],
        }),
      ],
    });

    //Create Table below CRF line
    const getChangeExecutionCheckbox = (label: string, selectedValue: string) => {
      return `${selectedValue === label ? '☑' : '☐'} ${label}`;
    };
    const getChangeCategoryCheckbox = (label: string, selectedValue: string) => {
      return `${selectedValue === label ? '☑' : '☐'} ${label}`;
    };
    const getNatureofChangeCheckbox = (label: string, selectedValue: string) => {
      return `${selectedValue === label ? '☑' : '☐'} ${label}`;
    };
    const checkbox = (checked: boolean) => (checked ? '☑' : '☐');

// Construct the paragraph with checkboxes
const typeOfChangeCell = new TableCell({
  columnSpan: 3, // merge if needed
  children: [
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_apicode)} API, Code Utilities, Libraries, Third-Party libraries etc.`,
          size: 20,
          // font: 'Arial Unicode MS', // optional but safe
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_architectural)} Architectural`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_audit_compliance)} Audit Compliance`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_configuration)} Configuration`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_design)} Design`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_hardware)} Hardware`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_network)} Network`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_oem_recommendation)} OEM Recommendation`,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_quality_changes)} Quality Changes (Common Services - logging, messages, etc.) `,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_Software)} Software Changes (i.e. OS, Database, Web, Application Server Software) `,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_security)} Security Level Check `,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_upgradation)} Upgradation (Hardware or Software) `,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${checkbox(this.formData.type_others)} Others (DB Impact) `,
          size: 20,
          // font: 'Arial Unicode MS',
        }),
      ],
    })
  ],
});
    const tablebelowCRFline = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'Project / Application Name',
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              columnSpan: 3,
              // width: { size: 75, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: (this.formData.application_name),
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: ('Initiation Date'),
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: (this.formData.initiation_date),
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'Requester Contact Number',
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: (this.formData.contact_number),
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'Change Execution',
                      bold: true,
                      size: 20
                    })
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeExecutionCheckbox('Normal', this.formData.change_execution), size: 20 }),

                  ]
                }),
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeExecutionCheckbox('Emergency', this.formData.change_execution), size: 20 }),

                  ]
                }),
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeExecutionCheckbox('Compliance', this.formData.change_execution), size: 20 }),

                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'Change Category',
                      bold: true,
                      size: 20
                    })
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeExecutionCheckbox('Major', this.formData.change_category), size: 20 }),
                  ]
                }),
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeCategoryCheckbox('Minor', this.formData.change_category), size: 20 }),
                  ]
                })
              ]
            }),

          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'Nature of Change',
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 75, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: getNatureofChangeCheckbox('Temporary', this.formData.nature_of_change), size: 20 }),
                  ]
                }),
                new Paragraph({
                  children: [
                    new TextRun({ text: getChangeCategoryCheckbox('Permanent', this.formData.nature_of_change), size: 20 }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: ('Implementation Date'),
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: (this.formData.implementation_date),
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
          ]
        }),
        new TableRow({
          children:[
            new TableCell({
              // width: { size: 25, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: ('Type of Change'),
                      bold: true,
                      size: 20,
                    }),
                  ]
                })
              ]
            }),
            typeOfChangeCell
          ]
        })
      ]
    });

    const ObjectiveofChangelabel= new Paragraph({
      spacing: {
        before: 300, // space before in twips (20 twips = 1 point)
      },
      children: [
        new TextRun({
          text: "Objective of Change: [Define objective of the change. This could include new business requirements, product feature enhancements or problem rectification.]",
          bold: true,
        }),
        new TextRun({ break: 2 })
    
      ],
    })

    const objectiveText: string = this.formData.ObjectiveofChange || "";

    // Split the content by newline
    const lines = objectiveText.split('\n');
    
    const ObjectiveofChangePara: Paragraph[] = [];
    
    lines.forEach((line) => {
      const trimmed = line.trim();
    
      if (trimmed.startsWith('•') || trimmed.startsWith('●')) {
        ObjectiveofChangePara.push(
          new Paragraph({
            text: trimmed.replace(/^•|^●/, '').trim(),
            bullet: { level: 0 },
          })
        );
      } else if (trimmed) {
        ObjectiveofChangePara.push(
          new Paragraph({
            children: [
              new TextRun({
                text: trimmed,
                break: 1,
              }),
            ],
          })
        );
      } else {
        // Empty line = paragraph break
        ObjectiveofChangePara.push(new Paragraph(""));
      }
    });



    const wordDoc = new Document({
      sections: [
        {
          headers: { default: header },
          footers: { default: footer },
          children: [
            titleAndCRFRow,
            new Paragraph({ spacing: { after: 300 } }), // space after title table
            tablebelowCRFline,
            ObjectiveofChangelabel,
            ...ObjectiveofChangePara   //Spread the array to flatten it or it will give error
            // new Paragraph(`Application Name: ${this.formData.application_name}`),
            // new Paragraph(`Description: ${this.formData.description}`),
            // new Paragraph(`Email: ${this.formData.email}`),
            // new Paragraph(`Details: ${this.formData.details}`),
          ],
        },
      ],
    });

    Packer.toBlob(wordDoc).then(blob => {
      saveAs(blob, "");
      this.router.navigate(['/change-request']);
    });
  }
}



