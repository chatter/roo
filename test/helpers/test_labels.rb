module TestLabels

  def expected_labels
    [
      ["anton", [[5, 3, "Sheet1"]]],
      ["berta", [[4, 2, "Sheet1"]]],
      ["caesar", [[7, 2, "Sheet1"]]],
      ["column_range",
        [ [5, 4, "Sheet1"],
          [6, 4, "Sheet1"],
          [7, 4, "Sheet1"],
          [8, 4, "Sheet1"],
          [9, 4, "Sheet1"]
        ]
      ],
      ["grid_range",
        [ [13, 2, "Sheet1"],
          [14, 2, "Sheet1"],
          [13, 3, "Sheet1"],
          [14, 3, "Sheet1"],
          [13, 4, "Sheet1"],
          [14, 4, "Sheet1"]
        ]
      ],
      ["non_adjacent_range",
        [ [11, 2, "Sheet1"],
          [11, 3, "Sheet1"],
          [11, 4, "Sheet1"],
          [13, 2, "Sheet1"],
          [14, 2, "Sheet1"],
          [13, 3, "Sheet1"],
          [14, 3, "Sheet1"],
          [13, 4, "Sheet1"],
          [14, 4, "Sheet1"],
          [5, 4, "Sheet1"],
          [6, 4, "Sheet1"],
          [7, 4, "Sheet1"],
          [8, 4, "Sheet1"],
          [9, 4, "Sheet1"]
        ]
      ],
      ["row_range",
        [ [11, 2, "Sheet1"],
          [11, 3, "Sheet1"],
          [11, 4, "Sheet1"]
        ]
      ]
    ]
  end

  def test_labels
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }


    with_each_spreadsheet(options) do |oo|
      assert_equal expected_labels, oo.labels, "error with labels array in class #{oo.class}"
    end
  end

  def test_labeled_cells
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      oo.default_sheet = oo.sheets.first
      begin
        row, col = oo.label("anton")[0]
      rescue ArgumentError
        puts "labels error at #{oo.class}"
        raise
      end
      assert_equal 5, row
      assert_equal 3, col

      row, col = oo.label("anton")[0]
      assert_equal "Anton", oo.cell(row, col)

      row, col = oo.label("berta")[0]
      assert_equal "Bertha", oo.cell(row, col)

      row, col = oo.label("caesar")[0]
      assert_equal "Cäsar", oo.cell(row, col)

      assert_empty oo.label("never")

      row, col = oo.label("column_range").last
      assert_equal "column_range_5", oo.cell(row, col)

      row, col = oo.label("grid_range")[3]
      assert_equal "grid_range_2_2", oo.cell(row, col)

      row, col = oo.label("non_adjacent_range")[5]
      assert_equal "grid_range_2_1", oo.cell(row, col)

      row, col = oo.label("row_range")[1]
      assert_equal "row_range_2", oo.cell(row, col)

      row, col, sheet = oo.label("anton")[0]
      assert_equal 5, row
      assert_equal 3, col
      assert_equal "Sheet1", sheet

      assert_equal "Anton", oo.anton
      assert_raises(NoMethodError) do
        row, col = oo.never
      end

      # Reihenfolge row, col,sheet analog zu #label
      assert_equal expected_labels, oo.labels, "error with labels array in class #{oo.class}"
    end
  end

  def test_label
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      begin
        row, col = oo.label("anton")[0]
      rescue ArgumentError
        puts "labels error at #{oo.class}"
        raise
      end

      assert_equal 5, row, "error with label in class #{oo.class}"
      assert_equal 3, col, "error with label in class #{oo.class}"

      row, col = oo.label("anton")[0]
      assert_equal "Anton", oo.cell(row, col), "error with label in class #{oo.class}"

      row, col = oo.label("berta")[0]
      assert_equal "Bertha", oo.cell(row, col), "error with label in class #{oo.class}"

      row, col = oo.label("caesar")[0]
      assert_equal "Cäsar", oo.cell(row, col), "error with label in class #{oo.class}"

      assert_empty oo.label("never")

      row, col = oo.label("column_range").last
      assert_equal "column_range_5", oo.cell(row, col)

      row, col = oo.label("grid_range")[3]
      assert_equal "grid_range_2_2", oo.cell(row, col)

      row, col = oo.label("non_adjacent_range")[5]
      assert_equal "grid_range_2_1", oo.cell(row, col)

      row, col = oo.label("row_range")[1]
      assert_equal "row_range_2", oo.cell(row, col)

      row, col, sheet = oo.label("anton")[0]
      assert_equal 5, row
      assert_equal 3, col
      assert_equal "Sheet1", sheet
    end
  end

  def test_method_missing_anton
    options = { name: "named_cells", format: [:openoffice, :excelx, :libreoffice] }
    with_each_spreadsheet(options) do |oo|
      # oo.default_sheet = oo.sheets.first
      assert_equal "Anton", oo.anton
      assert_raises(NoMethodError) do
        oo.never
      end
    end
  end
end
