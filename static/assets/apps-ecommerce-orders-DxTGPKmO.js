import "./app-DOPhicSe.js";
import { u as o } from "./apexcharts.esm-DPbJ6jlt.js";
class a {
  initOrdersOverview() {
    var r = {
        series: [
          {
            name: "Jumlah Case",
            data: jumlahCase,
          },
        ],
        chart: { height: 150, type: "bar", toolbar: { show: !1 } },
        plotOptions: {
          bar: { borderRadius: 5, dataLabels: { position: "top" } },
        },
        dataLabels: {
          enabled: !0,
          formatter: function (e) {
            return e;
          },
          style: { fontSize: "12px" },
        },
        grid: { padding: { bottom: -10 } },
        xaxis: {
          categories: mingguLabels,
          position: "bottom",
          axisBorder: { show: !1 },
          axisTicks: { show: !1 },
          tooltip: { enabled: !0 },
          title: {
            text: namaBulan,
            style: { fontSize: "18px", fontWeight: 600 },
          },
        },

        yaxis: {
          axisBorder: { show: !1 },
          axisTicks: { show: !1 },
          labels: { show: !0 },
        },
        colors: ["#2b7fff"],
      },
      t = new o(document.querySelector("#ordersOverview"), r);
    t.render();
  }
  init() {
    this.initOrdersOverview();
  }
}
document.addEventListener("DOMContentLoaded", function () {
  setTimeout(() => {
    new a().init();
  }, 100);
});
