self.onmessage = (e) => {
  const { records } = e.data;
  const XLSX = self.XLSX;

  const ws = XLSX.utils.json_to_sheet([], {
    header: ["序号","时间","类别","名称","规格","颜色","数量","备注","图片"]
  });

  records.forEach((r, i) => {
    XLSX.utils.sheet_add_json(ws, [{
      序号: i + 1,
      时间: r.time,
      类别: r.category,
      名称: r.name,
      规格: r.spec,
      颜色: r.color,
      数量: r.qty,
      备注: r.note,
      图片: r.image
    }], { skipHeader: true, origin: -1 });
  });

  ws["!images"] = records.map((r, i) => ({
    name: `IMG_${i+1}`,
    data: r.image.split(",")[1],
    opts: {
      base64: true,
      type: "jpeg",
      position: {
        type: "oneCellAnchor",
        from: { col: 8, row: i + 1 }
      },
      size: { width: 80, height: 80 }
    }
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "盘点数据");

  const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  self.postMessage(buf);
};
