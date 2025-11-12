import "./app-DOPhicSe.js";
import { u as o } from "./apexcharts.esm-DPbJ6jlt.js";
class c {
  initRadialChart() {
    var e = "",
      t = "",
      a = "",
      i = "";
    function p() {
      var l = "";
      if (((l = ["#2b7fff"]), l)) {
        var r = {
          series: [100],
          chart: {
            height: 110,
            type: "radialBar",
            parentHeightOffset: 0,
            sparkline: { enabled: !0 },
          },
          plotOptions: {
            radialBar: {
              hollow: { margin: 0, size: "50%" },
              track: { margin: 2 },
              dataLabels: { show: !1 },
            },
          },
          grid: { padding: { top: -15, bottom: -15 } },
          stroke: { lineCap: "round" },
          labels: ["Total Employee"],
          colors: l,
        };
        e != "" && e.destroy(),
          (e = new o(document.querySelector("#totalEmployee"), r)),
          e.render();
      }
      var n = "primary";
      if (((n = ["var(--color-purple-500)"]), n)) {
        var r = {
          series: [100],
          chart: { height: 110, type: "radialBar", sparkline: { enabled: !0 } },
          plotOptions: {
            radialBar: {
              hollow: { margin: 0, size: "50%" },
              track: { margin: 2 },
              dataLabels: { show: !1 },
            },
          },
          grid: { padding: { top: -15, bottom: -15 } },
          stroke: { lineCap: "round" },
          labels: ["Total Employee"],
          colors: n,
        };
        t != "" && t.destroy(),
          (t = new o(document.querySelector("#totalApplication"), r)),
          t.render();
      }
      var d = "";
      if (((d = ["#00c951"]), d)) {
        var r = {
          series: [100],
          chart: { height: 110, type: "radialBar", sparkline: { enabled: !0 } },
          plotOptions: {
            radialBar: {
              hollow: { margin: 0, size: "50%" },
              track: { margin: 2 },
              dataLabels: { show: !1 },
            },
          },
          grid: { padding: { top: -15, bottom: -15 } },
          stroke: { lineCap: "round" },
          labels: ["Total Employee"],
          colors: d,
        };
        a != "" && a.destroy(),
          (a = new o(document.querySelector("#hiredCandidates"), r)),
          a.render();
      }
      var s = "";
      if (((s = ["#ff6900"]), s)) {
        var r = {
          series: [100],
          chart: { height: 110, type: "radialBar", sparkline: { enabled: !0 } },
          plotOptions: {
            radialBar: {
              hollow: { margin: 0, size: "50%" },
              track: { margin: 2 },
              dataLabels: { show: !1 },
            },
          },
          grid: { padding: { top: -15, bottom: -15 } },
          stroke: { lineCap: "round" },
          labels: ["Total Employee"],
          colors: s,
        };
        i != "" && i.destroy(),
          (i = new o(document.querySelector("#rejectedCandidates"), r)),
          i.render();
      }
    }
    window.addEventListener("resize", function () {
      setTimeout(() => {
        p();
      }, 0);
    }),
      p();
  }
  initApplicationReceivedChart() {
  const fallback = {
    labels: ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
    series: [{
      name: "Tickets",
      type: "area",
      data: [0,0,0,0,0,0,0,0,0,0,0,0],
    }]
  };

  const injected = window.applicationReceived || fallback;

  const e = {
    series: injected.series,
    chart: { height: 315, type: "line", stacked: !1, margin: { left: 0, right: 0, top: 0, bottom: 0 }, toolbar: { show: !1 } },
    stroke: { width: [2], curve: "smooth" },
    fill: { opacity: [0.15], gradient: { inverseColors: !1, shade: "light", type: "vertical", opacityFrom: 0.85, opacityTo: 0.55, stops: [0, 100, 100, 100] } },
    labels: injected.labels,
    colors: ["#2b7fff"],
    markers: { size: 0 },
    grid: { padding: { top: -15, right: 0 } },
    tooltip: { shared: !0, intersect: !1, y: { formatter: a => typeof a < "u" ? a.toFixed(0) + " tickets" : a } },
  };

  const t = new o(document.querySelector("#applicationReceivedChart"), e);
  t.render();
}

  initTotalProjectChart() {
    var e = {
        series: [
          { name: "New", data: [44, 55, 41, 67, 22, 43, 14, 55, 41] },
          { name: "Pending", data: [13, 23, 20, 8, 13, 27, 8, 20, 8] },
          { name: "Completed", data: [11, 17, 15, 15, 21, 14, 24, 11, 17] },
          { name: "Rejected", data: [21, 7, 25, 13, 22, 8, 13, 7, 25] },
        ],
        chart: {
          type: "bar",
          height: 353,
          stacked: !0,
          zoom: { enabled: !0 },
          toolbar: { show: !1 },
          parentHeightOffset: 0,
        },
        plotOptions: {
          bar: { horizontal: !1, borderRadius: 2, columnWidth: "25%" },
        },
        grid: { padding: { top: -15, bottom: 5, right: 0 } },
        xaxis: {
          categories: ["01", "02", "03", "04", "05", "06", "07", "08", "09"],
        },
        dataLabels: { enabled: !1 },
        colors: ["#2b7fff", "#facc15", "#4ade80", "#f87171"],
        legend: { position: "bottom" },
        fill: { opacity: 1 },
      },
      t = new o(document.querySelector("#totalProjectChart"), e);
    t.render();
  }
  init() {
    this.initRadialChart(),
      this.initApplicationReceivedChart(),
      this.initTotalProjectChart();
  }
}
document.addEventListener("DOMContentLoaded", function () {
  setTimeout(() => {
    new c().init();
  }, 100);
});
