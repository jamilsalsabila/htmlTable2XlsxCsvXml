let args = process.argv.slice(2);
let pdfOutFile = "";
let csvOutFile = "";
let xmlOutFile = "";

const fs = require("fs");
const cheerio = require("cheerio");
const pdf = require("html-pdf");
const opts = {
  format: "A4",
  orientation: "landscape",
};

fs.open(args[0], "r", (err, fd) => {
  if (err) {
    console.log(err.message);
    return;
  }
  let buf = Buffer.alloc(1);
  fs.read(fd, buf, 0, 1, 0, (err, bytesRead, buf) => {
    if (err) {
      // console.log(err);
      fs.close(fd, (err) => {
        console.log(err.message);
        return;
      });
      return;
    }
    let $_ = cheerio();
    // console.log("bytes red: ", bytesRead);
    // console.log("buf:", buf);
    let s = buf.toString("utf-8");
    // console.log("s: ", s);
    if (s.charCodeAt(0) < 256) {
      console.log("Text file");
      fs.close(fd, (err) => {
        // console.log(err);
        return;
      });

      const data = fs.readFileSync(args[0], {
        encoding: "UTF-8",
      });
      const $ = cheerio.load(data);
      var html = `
<!DOCTYPE html>
	<head>
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
		<style>
			tr, td {
				border: 1px solid black;
				font-size: 12px;
				min-width: 100px;
			}
			table {
				border-collapse: collapse;
				margin-top:20px;
			}
		</style>
	</head>
	<body>
		<table>
		</table>
	</body>
	<script>
/*
 * HTML: Print Wide HTML Tables
 * https://salman-w.blogspot.com/2013/04/printing-wide-html-tables.html
 */
$(function() {
  // $("#print-button").on("click", function() {
    const width = 650;
    var table = $("table"),
      printWrap = $("<div></div>").insertAfter(table),
      i,
      printPage,
      start=0,
      tmp=[],
      c=0,
      tr = $('tr:last-child'),
      widthC=1;
    for(i = 1; i <= tr.children().length; ++i){
      if((tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width())>(width*widthC)){
        --i;
        tmp.push(start);
        tmp.push(tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width()-start);
        start=tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width();
        c++;c++;
        widthC++;
      }
    }
    if(tmp.length===0) {
      tmp.push(0);
      tmp.push(width);
      tmp[0]+=1;
    }
    else {
			tmp.push(start);
			tmp.push(tr.children(':last-child').position()['left']+tr.children(':last-child').width()-start);
			c+=2;
      tmp[1]-=3;
      for(i=2;i<tmp.length;i+=2){
        tmp[i]=-(tmp[i]-4);
        tmp[i+1]+=1;
      }
    }
    for(i=0;i<tmp.length;i+=2){
      printPage = $("<div></div>").css({
        "overflow": "hidden",
        "width": tmp[i+1],
        "page-break-before": i === 0 ? "auto" : "always"
      }).appendTo(printWrap);
      table.clone().removeAttr("id").appendTo(printPage).css({
        "position": "relative",
        "left": tmp[i]
      });
      $("<br />").appendTo(printPage);
    }
    printWrap.append("<p>Sumber: BPS Aceh</p>");
    table.remove();
    $(this).prop("disabled", true);
  // });
});
</script>
`;
      $_ = cheerio.load(html);
      $_("table").append($("#tableRightUp tbody").html());
      $_($("#tableLeftUp tbody tr").html()).insertBefore(
        "table tr:nth-child(1) th:nth-child(1)"
      );
      for (let i = 1; i <= $("#tableLeftBottom tbody tr").length; ++i) {
        $_("table").append(
          '<tr id="trth' +
          i +
          '">' +
          $("#tableLeftBottom tbody tr:nth-child(" + i + ")").html() +
          "</tr>"
        );
      }
      for (let i = 1; i <= $("table#tableRightBottom tr").length; ++i) {
        $_("table tr#trth" + i).append(
          $("#tableRightBottom tbody tr:nth-child(" + i + ")").html()
        );
      }
    } else {
      // console.log("Binary file");
      fs.close(fd, (err) => {
        // console.log(err);
        return;
      });
      let titleRow = 0;
      let totalColumns;
      let maxColumns = 0;
      let maxColumnsIndex = 0;
      let rowSpaces = [];
      let counter = 1;
      const XLSX = require("xlsx");
      let workbook = XLSX.readFile(args[0]);
      // console.log(XLSX.utils.sheet_to_html());
      if (args[1] === undefined) {
        args[1] = workbook.SheetNames[0];
      }
      // console.log(workbook.SheetNames[0]);
      $_ = cheerio.load(XLSX.utils.sheet_to_html(workbook.Sheets[args[1]]));
      for (let i = 1; i <= $_("table tr").length; ++i) {
        for (
          let j = $_("table tr:nth-child(" + i + ") td").length; j > 0;
          --j
        ) {
          if (
            $_(
              "table tr:nth-child(" + i + ") td:nth-child(" + j + ")"
            ).text() === ""
          ) {
            $_(
              "table tr:nth-child(" + i + ") td:nth-child(" + j + ")"
            ).remove();
          }
        }
        if ($_("table tr:nth-child(" + i + ")").text() === "") {
          // console.log($_("no content", "tr:nth-child(" + i + ")").text());
          rowSpaces.push(i);
        }
      }
      for (let i = rowSpaces.length - 1; i > -1; --i) {
        $_("table tr:nth-child(" + rowSpaces[i] + ")").remove();
      }
      // console.log($_.html());
      totalColumns = $_("table tr:last-child td").length;
      // console.log($_("table tr").length);

      while ($_("table tr:nth-child(" + (counter) + ") td").length === 1) {
        if ($_("table tr:nth-child(" + (counter) + ") td").text()) {
          $_("table tr:nth-child(" + (counter) + ") td").attr(
            "style",
            "border:none; word-wrap:break-word;"
          );
          $_("table tr:nth-child(" + (counter) + ") td").attr("rowspan", "1");
          $_("table tr:nth-child(" + (counter) + ")").attr("style", "border:none");
          titleRow++;
          counter++;
        }
      }

      for (let i = 1; i <= $_("table tr").length; ++i) {
        // console.log(i, $_("table tr:nth-child(" + i + ")").text());
        if ($_("table tr:nth-child(" + i + ") td").length > maxColumns) {
          maxColumns = $_("table tr:nth-child(" + i + ") td").length;
          maxColumnsIndex = i;
        }
      }
      if (titleRow != 0) {
        for (let i = 1; i <= titleRow; ++i) {
          $_("table tr:nth-child(" + i + ") td").attr("colspan", maxColumns);
        }
      }
      $_("head").append(`
			<style>
					tr, td {
						border:1px solid black;
						font-size: 12px;
						min-width: 100px;
					}
					table {
						//margin: 50px;
						border-collapse:collapse;
					}
					#th2b{
						background-color: rgb(0, 33, 66); color: white; height: 64px; min-width: 201px; width: 201px; vertical-align: middle; text-align:center;
					}
					#th1{
						background-color: rgb(0, 33, 66); color: white; width: 204px; height: 12px; vertical-align: middle; text-align: center;
					}
					#th2{
						background-color: rgb(0, 33, 66); color: white; width: 100px; height: 12px; vertical-align: middle; text-align: center;
					}
					#th4odd{
						height: 24px; width: 201px; vertical-align: middle;
					}
					#th4even{
						background-color: #d4d4d4; color: black; height: 24px; width: 201px; vertical-align: middle;
					}
					.datasodd{
						text-align: right; height: 24px; width: 90px; vertical-align: middle;
					}
					.dataseven{
						background-color: #d4d4d4; color: black; text-align: right; height: 24px; width: 90px; vertical-align: middle;
					}
				</style>
			`);

      let c;
      let t = titleRow + 1;
      let thRowSpan = parseInt(
        $_(`table tr:nth-child(${t}) td:nth-child(1)`).attr("rowspan") == null ?
        1 :
        $_(`table tr:nth-child(${t}) td:nth-child(1)`).attr("rowspan")
      );
      let rowSpan = thRowSpan;
      // console.log(thRowSpan, maxColumnsIndex, titleRow, maxColumns, t);

      $_("table tr:nth-child(" + t + ") td:nth-child(1)").attr("id", "th2b");
      if (thRowSpan === 1) {
        for (c = 2; c <= $_("table tr:nth-child(" + t + ") td").length; c++) {
          $_("table tr:nth-child(" + t + ") td:nth-child(" + c + ")").attr(
            "id",
            "th2"
          );
        }
        // maxColumnsIndex++;
        t++;
      } else {
        rowSpan -=
          $_("table tr:nth-child(" + t + ") td:nth-child(" + 2 + ")").attr(
            "rowspan"
          ) == null ?
          1 :
          $_("table tr:nth-child(" + t + ") td:nth-child(" + 2 + ")").attr(
            "rowspan"
          );
        for (c = 2; c <= $_("table tr:nth-child(" + t + ") td").length; c++) {
          $_("table tr:nth-child(" + t + ") td:nth-child(" + c + ")").attr(
            "id",
            "th1"
          );
        }
        if (rowSpan !== 0) {
          t++;
          for (; t <= titleRow + thRowSpan; t++) {
            if (rowSpan === 1) {
              for (
                c = 1; c <= $_("table tr:nth-child(" + t + ") td").length; c++
              ) {
                $_(
                  "table tr:nth-child(" + t + ") td:nth-child(" + c + ")"
                ).attr("id", "th2");
              }
            } else if (rowSpan > 1) {
              for (
                c = 1; c <= $_("table tr:nth-child(" + t + ") td").length; c++
              ) {
                $_(
                  "table tr:nth-child(" + t + ") td:nth-child(" + c + ")"
                ).attr("id", "th1");
              }
            }
            rowSpan -=
              $_("table tr:nth-child(" + t + ") td:nth-child(" + 1 + ")").attr(
                "rowspan"
              ) == null ?
              1 :
              $_(
                "table tr:nth-child(" + t + ") td:nth-child(" + 1 + ")"
              ).attr("rowspan");
          }
        } else {
          for (c = 1; c <= $_("table tr:nth-child(" + t + ") td").length; c++) {
            $_("table tr:nth-child(" + t + ") td:nth-child(" + c + ")").attr(
              "rowspan",
              1
            );
          }
          t++;
        }
      }
      // console.log($_.html());
      for (; t <= $_("table tr").length; ++t) {
        if ($_("table tr:nth-child(" + t + ") td").length === maxColumns) {
          if (t % 2 === 1) {
            $_("table tr:nth-child(" + t + ") td:nth-child(1)").attr(
              "id",
              "th4odd"
            );
            for (
              let i = 2; i <= $_("table tr:nth-child(" + t + ") td").length;
              ++i
            ) {
              $_("table tr:nth-child(" + t + ") td:nth-child(" + i + ")").attr(
                "class",
                "datasodd"
              );
            }
          } else {
            $_("table tr:nth-child(" + t + ") td:nth-child(1)").attr(
              "id",
              "th4even"
            );
            for (
              let i = 2; i <= $_("table tr:nth-child(" + t + ") td").length;
              ++i
            ) {
              $_("table tr:nth-child(" + t + ") td:nth-child(" + i + ")").attr(
                "class",
                "dataseven"
              );
            }
          }
        } else if ($_("table tr:nth-child(" + t + ") td").length === 1) {
          $_("table tr:nth-child(" + t + ") td:last-child").css({
            border: "none",
          });
          $_("table tr:nth-child(" + t + ")").css({
            border: "none",
          });
          $_("table tr:nth-child(" + t + ") td:last-child").attr(
            "colspan",
            maxColumns
          );
        }
      }
      // console.log($_.html());
      $_("html").append(`
			<script>
			/*
			 * HTML: Print Wide HTML Tables
			 * https://salman-w.blogspot.com/2013/04/printing-wide-html-tables.html
			 */
			$(function() {
				// $("#print-button").on("click", function() {
					const width = 650;
					var table = $("table"),
						printWrap = $("<div></div>").insertAfter(table),
						i,
						printPage,
						start=0,
						tmp=[],
						c=0,
						tr = $('tr:nth-child(${maxColumnsIndex})'),
						widthC=1;
					for(i = 1; i <= tr.children().length; ++i){
						if((tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width())>(width*widthC)){
							--i;
							tmp.push(start);
							tmp.push(tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width()-start);
							start=tr.children(':nth-child('+i+')').position()['left']+tr.children(':nth-child('+i+')').width();
							c++;c++;
							widthC++;
						}
					}
					if(tmp.length===0) {
						tmp.push(0);
						tmp.push(width);
						tmp[0]+=1;
					}
					else {
						tmp.push(start);
						tmp.push(tr.children(':last-child').position()['left']+tr.children(':last-child').width()-start);
						c+=2;
						tmp[1]-=3;
						for(i=2;i<tmp.length;i+=2){
							tmp[i]=-(tmp[i]-4);
							tmp[i+1]+=1;
						}
					}
					for(i=0;i<tmp.length;i+=2){
						printPage = $("<div></div>").css({
							"overflow": "hidden",
							"width": tmp[i+1],
							"page-break-before": i === 0 ? "auto" : "always"
						}).appendTo(printWrap);
						table.clone().removeAttr("id").appendTo(printPage).css({
							"position": "relative",
							"left": tmp[i]
						});
						$("<br />").appendTo(printPage);
					}
					printWrap.append("<p>Sumber: BPS Aceh</p>");
					table.remove();
					$(this).prop("disabled", true);
				// });
			});
			</script>`);
      $_("head").append(`
			<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
			`);
    }
    if (args[1] === undefined) {
      pdfOutFile = args[0].replace(/xls/g, "pdf");
      csvOutFile = args[0].replace(/xls/g, "csv");
      xmlOutFile = args[0].replace(/xls/g, "xml")
    } else {
      pdfOutFile = args[0].replace(/xls/g, "pdf").replace(/\.pdf$/, `_${args[1]}.pdf`).replace(/\s+/, "_");
      csvOutFile = args[0].replace(/xls/g, "csv").replace(/\.csv$/, `_${args[1]}.csv`).replace(/\s+/, "_");
      xmlOutFile = args[0].replace(/xls/g, "xml").replace(/\.xml$/, `_${args[1]}.xml`).replace(/\s+/, "_");
    }
    pdf
      .create($_.html(), opts)
      .toFile(pdfOutFile, function (err, res) {
        if (err) return console.log(err);
        // console.log(res);
      });
    console.log($_.html());
    // [ CSV ] //
    // console.log($_.html());
    let header = [];
    let body = [];
    let tmp = [];
    let maxHeaderRow = $_("#th2b").attr("rowspan");
    let child = cheerio();
    let colspan = 0;
    let k = 1;
    let firstRow = cheerio();
    let data = "";
    if ($_("#th2b").html() === null) {
      child = $_("table tr");
      for (let c = 1; c <= child.length; ++c) {
        header.push(child.children(":nth-child(" + 1 + ")").html().replace(/[\n\s]+/, " ")
          .trim());
        child = child.next();
      }
      data += header.join("\t");
    } else {
      header.push(
        $_("#th2b")
        .html()
        .replace(/[\n\s]+/, " ")
        .trim()
      );
      firstRow = $_("#th2b").parent();
      firstRow.children().each((index, elem) => {
        if (index == 1) {
          maxHeaderRow -=
            elem.attribs["rowspan"] === undefined ? 1 : elem.attribs["rowspan"];
        }
        if (index > 0) {
          if (!elem.attribs["colspan"]) {
            header.push(elem.firstChild.data);
          } else {
            // console.log(index, parseInt(elem.attribs["colspan"]));
            for (let i = 0; i < parseInt(elem.attribs["colspan"]); ++i) {
              header.push(elem.firstChild.data);
            }
          }
        }
      });
      while (maxHeaderRow > 0) {
        firstRow = firstRow.next();
        maxHeaderRow -=
          firstRow.children(":nth-child(1)").attr("rowspan") === undefined ?
          1 :
          firstRow.children(":nth-child(1)").attr("rowspan");
        // console.log(maxHeaderRow);
        for (let i = 1; i <= firstRow.children().length; ++i) {
          child = firstRow.children(":nth-child(" + i + ")");
          // console.log(child.html());
          colspan = child.attr("colspan") === undefined ? 1 : child.attr("colspan");
          // console.log(colspan);
          if (colspan > 1) {
            for (let j = 0; j < colspan; ++j) {
              header[k++] += `->${child
								.html()
								.replace(/[\n\s\t]+/, " ")
								.trim()}`;
            }
          } else {
            header[k++] += `->${child
							.html()
							.replace(/[\n\s\t]+/, " ")
							.trim()}`;
          }
        }
        k = 1

      }
      // console.log(header);
      firstRow = firstRow.next();
      while (firstRow.html() !== null) {
        // console.log(firstRow.html());
        // console.log(firstRow.html() !== null);
        if (firstRow.children().length > 1) {
          for (let i = 1; i <= firstRow.children().length; ++i) {
            if (firstRow.children(":nth-child(" + i + ")").attr("colspan") > 1) {
              for (let z = 0; z < parseInt(firstRow.children(":nth-child(" + i + ")").attr("colspan")) - 1; ++z) {
                tmp.push(" ");
              }
            }
            tmp.push(
              firstRow
              .children(":nth-child(" + i + ")")
              .html()
              .replace(/[\n\s\t]+/, " ")
              .trim()
            );
          }
          body.push(tmp);
          tmp = [];
        }
        firstRow = firstRow.next();
      }
      data += header.join("\t");
      data += "\n";
      for (let i = 0; i < body.length; ++i) {
        data += body[i].join("\t");
        data += "\n";
      }
    }
    // console.log(data);
    fs.writeFileSync(csvOutFile, data);
    // console.log(firstRow.html());

    // [ XML ] //
    tmp = [];
    header = [];
    body = [];
    colspan = "";
    maxHeaderRow = parseInt($_("#th2b").attr("rowspan"));
    firstRow = $_("#th2b").parent();
    let rowspan = "";
    let childText = "";

    maxHeaderRow -= (firstRow.children(":nth-child(" + 2 + ")").attr("rowspan") === undefined) ? 1 : parseInt(firstRow.children(":nth-child(" + 2 + ")").attr("rowspan"));
    for (let i = 1; i <= firstRow.children().length; ++i) {
      child = firstRow.children(":nth-child(" + i + ")");
      childText = child.html().replace(/[\s]+/g, " ").trim();
      rowspan = (child.attr("rowspan") === undefined) ? '1' : child.attr("rowspan");
      colspan = (child.attr("colspan") === undefined) ? '1' : child.attr("colspan");
      tmp.push(childText, rowspan, colspan);
    }
    header.push(tmp);
    tmp = [];
    // console.log(header);

    firstRow = firstRow.next();
    while (maxHeaderRow > 0) {
      maxHeaderRow -=
        firstRow.children(":nth-child(1)").attr("rowspan") === undefined ?
        1 :
        parseInt(firstRow.children(":nth-child(1)").attr("rowspan"));

      for (let i = 1; i <= firstRow.children().length; ++i) {
        child = firstRow.children(":nth-child(" + i + ")");
        childText = child.html().replace(/[\s]+/g, " ").trim();
        rowspan = (child.attr("rowspan") === undefined) ? '1' : child.attr("rowspan");
        colspan = (child.attr("colspan") === undefined) ? '1' : child.attr("colspan");
        tmp.push(childText, rowspan, colspan);
      }
      header.push(tmp);
      tmp = [];
      firstRow = firstRow.next();
    }
    // console.log(header);
    tmp = [];
    // firstRow = firstRow.next();
    while (firstRow.html() !== null) {
      if (firstRow.children().length > 1) {
        for (let i = 1; i <= firstRow.children().length; ++i) {
          child = firstRow.children(":nth-child(" + i + ")");
          childText = child.html().replace(/[\s]+/g, " ").trim();
          rowspan = (child.attr("rowspan") === undefined) ? '1' : child.attr("rowspan");
          colspan = (child.attr("colspan") === undefined) ? '1' : child.attr("colspan");
          tmp.push(childText, rowspan, colspan);
        }
        body.push(tmp);
        tmp = [];
      }
      firstRow = firstRow.next();
    }
    // console.log(body);
    let headerLength = header.length;
    let bodyLength = body.length;
    const pug = require('pug');
    const templateFunc = pug.compileFile("./main_2.pug");
    fs.writeFileSync(xmlOutFile, templateFunc({
      descr: "bla bla",
      headerLength: headerLength,
      header: header,
      bodyLength: bodyLength,
      body: body,
    }));
  });
  fs.close(fd, (err) => {
    // console.log(err);
    return;
  });
});