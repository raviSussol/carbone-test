const { writeFileSync } = require("fs");
const { render } = require("carbone");

// Data to inject
const data = {
  lmaip: {
    UNFPA_deliveries_summary: {
      data: [
        {
          item_name: "2-MET Plasma ELISA",
          units_received: 28,
          value_received: 0,
        },
        {
          item_name: "AccuDiag VIsE1/pepC10 IgG/IgM ELISA",
          units_received: 5,
          value_received: 0,
        },
        {
          item_name: "ACTICLOT dPT",
          units_received: 2,
          value_received: 0,
        },
      ],
    },
    CW_dist_summary: {
      data: [
        {
          store_name: "General",
        },
      ],
    },
    usdRate: 0.74,
    reportEndDate: "2021-12-28",
  },
};

const generateReport = () => {
  // Generate a report using the sample template provided by carbone module
  // This LibreOffice template contains "Hello {d.firstname} {d.lastname} !"
  // Of course, you can create your own templates!
  render(
    "./sample.ods",
    data,
    { convertTo: "xlsx", lang: "en-us" },
    function (err, result) {
      if (err) {
        return console.log(err);
      }
      // write the result
      writeFileSync("result.xlsx", result);
      process.exit(); // To kill automatically LibreOffice workers
    }
  );
};
generateReport();
