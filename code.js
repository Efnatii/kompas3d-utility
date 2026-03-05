const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, 
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, BorderStyle, 
        WidthType, TabStopType, TabStopPosition, UnderlineType, ShadingType, 
        VerticalAlign, PageNumber, PageBreak } = require('docx');
const fs = require('fs');
const path = require('path');

// Helper function to create text run with specific formatting
function textRun(text, options = {}) {
  return new TextRun({
    text: text,
    font: options.font || "Times New Roman",
    size: options.size || 24, // 12pt = 24 half-points
    bold: options.bold || false,
    italics: options.italics || false,
    ...options
  });
}

// Create signature block with stamp image
function createSignatureBlock(imagePath, name, role) {
  const children = [
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 240, after: 120 },
      children: [textRun(role + ":", { bold: true })]
    })
  ];
  
  // Add stamp image if available
  if (fs.existsSync(imagePath)) {
    try {
      const imageBuffer = fs.readFileSync(imagePath);
      children.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 60 },
          children: [
            new ImageRun({
              type: "jpg",
              data: imageBuffer,
              transformation: { width: 120, height: 120 },
              altText: { title: "Печать", description: "Печать и подпись", name: "stamp" }
            })
          ]
        })
      );
    } catch (e) {
      console.log("Could not add image:", e.message);
    }
  }
  
  children.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 60 },
      children: [textRun("_________________ /" + name + "/")]
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 60 },
      children: [textRun("                             (подпись)")]
    }),
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { after: 120 },
      children: [textRun("М.П.")]
    })
  );
  
  return children;
}

// Create the document
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Times New Roman", size: 24 }
      }
    },
    paragraphStyles: [
      {
        id: "Title",
        name: "Title",
        basedOn: "Normal",
        run: { size: 32, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 0, after: 120 }, alignment: AlignmentType.CENTER }
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        run: { size: 26, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 240, after: 120 } }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        run: { size: 24, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 120, after: 60 } }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 1134, right: 850, bottom: 1134, left: 1701 }, // Standard Russian document margins
        size: { width: 11906, height: 16838 } // A4
      }
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 20 })]
          })
        ]
      })
    },
    children: [
      // ========== PAGE 1 ==========
      // Title
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [textRun("ДОГОВОР ПОДРЯДА", { bold: true, size: 28 })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 120 },
        children: [textRun("№ 22-10-18", { size: 24 })]
      }),
      
      // City and date on same line
      new Paragraph({
        spacing: { after: 240 },
        children: [
          textRun("г. Санкт-Петербург"),
          new TextRun({ text: "\t" }),
          textRun("«21» октября 2022 г.")
        ]
      }),
      
      // Introduction paragraph
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("Общество с ограниченной ответственностью «Производственно-коммерческая фирма «Форвард» (далее - ООО «ПКФ «Форвард»), именуемое в дальнейшем Заказчик, в лице Генерального директора Холодилиной Александры Романовны, действующего на основании Устава, с одной стороны, и")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("Общество с ограниченной ответственностью «ИНКО» (ООО «ИНКО»), именуемое в дальнейшем Подрядчик, в лице Генерального директора Витьоненко Александра Владиславовича, действующего на основании Устава, с другой стороны,")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 240 },
        indent: { firstLine: 720 },
        children: [
          textRun("далее совместно именуемые – «Стороны», заключили настоящий договор подряда (далее по тексту – «Договор») о нижеследующем:")
        ]
      }),
      
      // Section 1
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("1. Предмет Договора", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("1.1. Заказчик поручает, а Подрядчик принимает на себя выполнение комплекса работ по проектированию систем СПС, СОУЭ, ОС, поставке оборудования и материалов, монтажным и пуско-наладочным работам (далее – Работы) для объекта: Блочно-модульное здание (БМЗ) для размещения оборудования Комплектной трансформаторной подстанции (КТП) разработки и производства Заказчика, (далее – Объект)")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("1.2. Заказчик обязуется принять и оплатить надлежащим образом выполненные работы.")
        ]
      }),
      
      // Section 2
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("2. Сроки действия Договора", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("2.1. Настоящий Договор действует с момента его подписания Сторонами до полного исполнения ими обязательств по Договору и завершения всех взаиморасчетов.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("2.2. Срок выполнения работ:")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 0 },
        indent: { firstLine: 720 },
        children: [
          textRun("начало работ – с момента подписания настоящего Договора;")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("окончание работ – в срок по 24.10.2022 г.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("2.3. Сроки выполнения работ могут изменяться по согласованию Сторон.")
        ]
      }),
      
      // Section 3
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("3. Стоимость работ и порядок расчетов", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.1. Общая стоимость работ по настоящему Договору в соответствии с Протоколом согласования о договорной цене (Приложение № 1) составляет 120 000 (Сто двадцать тысяч) рублей 00 копеек, в том числе НДС 20% – 20 000 руб. 00 коп.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("Цена Договора является твердой и не подлежит изменению в течение срока действия Договора.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.2. Общая стоимость работ по настоящему Договору оплачивается Заказчиком в следующем порядке:")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.2.1. Заказчик производит авансовый платеж в размере 100% от стоимости оборудования и материалов и проектных работ, согласно Приложению №1 к настоящему Договору, что составляет 85 000 (Восемьдесят пять тысяч) рублей 00 копеек, в том числе НДС 20% –14 166 руб. 67 коп., в течение 3 (трех) банковских дней с момента подписания настоящего Договора.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.2.2. Заказчик производит окончательную оплату выполненных работ в размере 100% от стоимости монтажных и пусконаладочных работ, согласно Приложению №1 к настоящему Договору, что составляет 35 000 (Тридцать пять тысяч) рублей 00 копеек, в том числе НДС 20% – 5 833 руб. 33 коп., в течение 3 (трех) банковских дней от даты подписания обеими Сторонами соответствующего Акта сдачи-приемки выполненных работ по Договору.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.3. Оплата по настоящему Договору осуществляется путем перечисления денежных средств с расчетного счета Заказчика на расчетный счет Подрядчика. Моментом осуществления оплаты является момент списания денежных средств с корреспондентского счета банка Заказчика.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.4. В случае изменения объема работ, стоимость работ подлежит корректировке и будет отражена в дополнительных соглашениях.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("3.5. Стоимость работ, порядок и сроки их оплаты, могут быть изменены, не иначе как на основании дополнительных соглашений к настоящему Договору, подписываемых обеими Сторонами.")
        ]
      }),
      
      // Section 4
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("4. Порядок сдачи и приемки работ", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.1. Заказчик назначает на объекте ответственных представителей, которые от имени Заказчика в пределах своих полномочий будут осуществлять технический надзор и контроль за выполнением работ, производить проверку качества работ.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.2. Выполненные работы принимаются комиссией, в которую входят представители Заказчика и Подрядчика.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.3. Комиссия составляет акт по каждому объекту, в котором подтверждается факт выполнения работ или указываются недостатки и сроки их устранения.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.4. Заказчик в течение 5 (Пяти) рабочих дней со дня получения акта сдачи-приемки выполненных работ и отчетных документов по каждому объекту обязан направить Подрядчику подписанный акт или мотивированный отказ от приемки работ.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.5. В случае мотивированного отказа от приемки работ, Сторонами составляется двусторонний акт с перечнем необходимых доработок и сроков их выполнения.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("4.6. Если в указанный срок от Заказчика не поступил мотивированный отказ, то работа считается выполненной и принятой Заказчиком.")
        ]
      }),
      
      // Section 5
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("5. Ответственность Сторон", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.1. Заказчик обязуется:")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.1.1. Передать Подрядчику документацию, необходимую для выполнения работ по настоящему Договору.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.1.2. Обеспечить строительную готовность Объекта, бытовое помещение для рабочих и охраняемое помещение для хранения оборудования, материалов, инструментов с момента вступления настоящего Договора в силу, обеспечить свободный доступ к объектам в рабочее время, другое время по взаимному согласию")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2. Подрядчик обязуется:")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2.1. Выполнить все работы в объеме п. 1.1 и в сроки, предусмотренные разделом 2 настоящего Договора, в полном соответствии с действующими нормативно-техническими документами, строительными нормами и правилами.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2.2. Нести ответственность в размере причиненного ущерба перед Заказчиком за ненадлежащее исполнение обязательств по настоящему Договору.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2.3. Неукоснительно выполнять на объекте Заказчика все необходимые меры по технике безопасности, пожарной безопасности и меры по охране окружающей среды.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2.4. Соблюдать правила пропускного и внутриобъектного режима, действующего у Заказчика.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("5.2.5. После окончания работ вернуть по акту всю документацию Заказчика, переданную Подрядчику согласно п. 5.1.1.")
        ]
      }),
      
      // Section 6
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("6. Гарантии", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("6.1. Подрядчик гарантирует соответствие оказываемых Услуг/выполняемых Работ требованиям, предъявляемым законодательством РФ, СНИП, техническими нормами к услугам такого рода, и требованиям Договора и отсутствие обстоятельств, препятствующих достижению желаемого Заказчиком результата в отношении Услуг/Работ, оказываемых по Договору.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("6.2. Гарантийный срок в отношении результатов оказанных Услуг/выполненных Работ составляет 12 (Двенадцать) месяцев с даты подписания сторонами акта сдачи – приемки выполненных по договору работ в полном объеме.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("6.3. В течение гарантийного срока Подрядчик за свой счет устраняет все обнаруженные и выявленные недостатки выполненных Работ по настоящему Договору в согласованный Сторонами срок (но не более 15 рабочих дней).")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("6.4. В случае не устранения Подрядчиком недостатков в установленный срок, в порядке, определенном пунктом 6.3. Договора, Заказчик имеет право устранить недостатки своими силами или силами третьих лиц с отнесением расходов на Подрядчика. Подрядчик в таком случае обязан компенсировать Заказчику расходы по устранению недостатков в течение 5 (пяти) рабочих дней с даты получения соответствующего счета на оплату от Заказчика.")
        ]
      }),
      
      // Section 7
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("7. Условия расторжения Договора", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("7.1. Досрочное расторжение Договора возможно по взаимному согласию Сторон.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("7.2. В случае невыполнения обязательств одной из Сторон другая Сторона имеет право расторгнуть Договор досрочно, предупредив в десятидневный срок другую Сторону. При этом, нарушившая свои обязательства Сторона не имеет права на компенсацию произведенных затрат и убытков. Во всех иных случаях расторжение Договора в одностороннем порядке не допускается. Сторона, нарушившая это условие, обязана возместить другой стороне все убытки, включая понесенные расходы к моменту расторжения Договора.")
        ]
      }),
      
      // Section 8
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("8. Заключительные условия", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("8.1. Во всех случаях, не предусмотренных настоящим Договором, Стороны руководствуются действующим гражданским законодательством.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("8.2. Все споры между Сторонами решаются путем переговоров, а в случае принципиальных разногласий – в Арбитражном суде. Спор, возникающий из настоящего Договора, может быть передан на разрешение Арбитражного суда города Санкт-Петербурга и Ленинградской области после принятия Сторонами мер по досудебному урегулированию по истечении тридцати календарных дней со дня направления претензии (требования).")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("8.3. Договор имеет силу, если при его заключении Стороны подписывают соответствующие документы путем обмена сообщениями по электронной почте, каждое из которых содержит копию документа, подписанную уполномоченным представителем направившей сообщение Стороны.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("8.4. Наименования разделов Договора приведены исключительно для удобства, не имеют юридической силы и не влияют на размещение пунктов в Договоре.")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("Настоящий Договор составлен в 2 (двух) идентичных экземплярах на русском языке, имеющих равную юридическую силу, по одному экземпляру для каждой Стороны.")
        ]
      }),
      
      // Section 9
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("9. Приложения к Договору", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("9.1. Приложение № 1 – Протокол согласования о договорной цене.")
        ]
      }),
      
      // Section 10
      new Paragraph({
        spacing: { before: 240, after: 120 },
        children: [textRun("10. Юридические адреса и банковские реквизиты Сторон", { bold: true, size: 24 })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 60 },
        children: [
          textRun("10.1 Заказчик: ", { bold: true }),
          textRun("ООО «ПКФ «Форвард»")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("Юридический адрес: 192171, г. Санкт-Петербург, ул. Седова, д.57")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("Почтовый адрес: 192171, г. Санкт-Петербург, ул. Седова, д.57, литер В, помещение 7-Н/15")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("ИНН 7811744437, КПП 781101001, ОГРН 1207800017314, ОКПО 43392559")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("р/счет 40702810600000000778 Филиал «Центральный» банка ВТБ (ПАО) г. Москва")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("К/с 30101810145250000411, БИК 044525411")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 120 },
        children: [textRun("Тел.: +7 (812) 372-58-80, e-mail: director@forward-ltd.ru, info@forward-ltd.ru")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 60 },
        children: [
          textRun("10.2. Подрядчик: ", { bold: true }),
          textRun("ООО «ИНКО»")
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("Юридический адрес: 197022, г. Санкт-Петербург, пр-т Аптекарский, д. 6, литер А, пом. 6-Н, оф. 603")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("Почтовый адрес: 197022, г. Санкт-Петербург, пр-т Аптекарский, д. 6, литер А, пом. 6-Н, оф. 603")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("ИНН 7813638040, КПП 781301001, ОГРН 1197847175327, ОКПО 41297963")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("Р/с 40702810055000052102 СЕВЕРО-ЗАПАДНЫЙ БАНК ПАО СБЕРБАНК г. Санкт-Петербург")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 0 },
        children: [textRun("К/с 30101810500000000653, БИК 044030653")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 250, after: 240 },
        children: [textRun("Тел. +7 (812) 448-39-01, e-mail: oooinko@internet.ru")]
      }),
      
      // Signatures section with stamps
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 240, after: 120 },
        children: [textRun("От Заказчика:", { bold: true })]
      }),
      // Add signature page image (page 3 has the signatures with stamps)
      ...(fs.existsSync("/home/z/my-project/upload/ООО ИНКО_page-0003.jpg") ? [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 60 },
          children: [
            new ImageRun({
              type: "jpg",
              data: fs.readFileSync("/home/z/my-project/upload/ООО ИНКО_page-0003.jpg"),
              transformation: { width: 200, height: 280 },
              altText: { title: "Страница с подписями", description: "Подписи и печати сторон", name: "signatures" }
            })
          ]
        })
      ] : []),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("_________________ /Холодилина А.Р./")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("                             (подпись)")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
        children: [textRun("М.П.")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 240, after: 120 },
        children: [textRun("От Подрядчика:", { bold: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("_________________ /Вивтоненко А.В./")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("                             (подпись)")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
        children: [textRun("М.П.")]
      }),
      
      // Page break for Appendix
      new Paragraph({ children: [new PageBreak()] }),
      
      // ========== PAGE 4 - APPENDIX ==========
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 60 },
        children: [textRun("Приложение №1", { bold: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 240 },
        children: [textRun("к Договору подряда № 22-10-18 от «21» октября 2022 г.")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [textRun("ПРОТОКОЛ", { bold: true, size: 28 })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [textRun("соглашения о договорной цене")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [textRun("на выполнение комплекса работ")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [textRun("по Договору подряда № 22-10-18 от «21» октября 2022 г.")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { firstLine: 720 },
        children: [
          textRun("Мы, нижеподписавшиеся, от лица Заказчика - Генеральный директор ООО «ПКФ «Форвард» Холодилина Александра Романовна, с одной стороны, и от лица Подрядчика - Генеральный директор ООО «ИНКО» Вивтоненко Александр Владиславович, с другой стороны, удостоверяем в том, что Сторонами достигнуто соглашение о величине договорной цены на выполнение комплекса работ по проектированию систем СПС, СОУЭ, ОС, поставке оборудования и материалов, монтажным и пуско-наладочным для объекта: Блочно-модульное здание (БМЗ) для размещения оборудования Комплектной трансформаторной подстанции (КТП) разработки и производства Заказчика, в сумме:")
        ]
      }),
      
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 240 },
        children: [textRun("120 000 (Сто двадцать тысяч) рублей 00 копеек, в том числе НДС 20% – 20 000 руб. 00 коп.", { bold: true })]
      }),
      
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("В том числе:", { bold: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 60 },
        indent: { left: 360 },
        children: [textRun("1. Оборудование и материалы – 65 000 руб. 00 коп.")]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 60 },
        indent: { left: 360 },
        children: [textRun("2. Монтажные работы – 30 000 руб. 00 коп.")]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 60 },
        indent: { left: 360 },
        children: [textRun("3. Пуско-наладочные работы – 5 000 руб. 00 коп.")]
      }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 120 },
        indent: { left: 360 },
        children: [textRun("4. Рабочая и исполнительная документация – 20 000 руб. 00 коп.")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 250, after: 240 },
        indent: { firstLine: 720 },
        children: [
          textRun("Настоящий протокол является основанием для проведения взаимных расчетов и платежей между Заказчиком и Подрядчиком.")
        ]
      }),
      
      // Appendix signatures with page 4 image
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 240, after: 120 },
        children: [textRun("От Заказчика:", { bold: true })]
      }),
      // Add appendix signature page image
      ...(fs.existsSync("/home/z/my-project/upload/ООО ИНКО_page-0004.jpg") ? [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { after: 60 },
          children: [
            new ImageRun({
              type: "jpg",
              data: fs.readFileSync("/home/z/my-project/upload/ООО ИНКО_page-0004.jpg"),
              transformation: { width: 200, height: 280 },
              altText: { title: "Приложение с подписями", description: "Подписи и печати приложения", name: "appendix-signatures" }
            })
          ]
        })
      ] : []),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("_________________ /Холодилина А.Р./")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("                             (подпись)")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
        children: [textRun("М.П.")]
      }),
      
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { before: 240, after: 120 },
        children: [textRun("От Подрядчика:", { bold: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("_________________ /Вивтоненко А.В./")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 60 },
        children: [textRun("                             (подпись)")]
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { after: 120 },
        children: [textRun("М.П.")]
      })
    ]
  }]
});

// Save the document
Packer.toBuffer(doc).then(buffer => {
  const outputPath = path.join(__dirname, "Договор_подряда_ООО_ИНКО.docx");
  fs.writeFileSync(outputPath, buffer);
  console.log("Document created successfully!");
});
