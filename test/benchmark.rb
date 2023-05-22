#!/usr/bin/env ruby -s
# frozen_string_literal: true

$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"
require 'axlsx'
require 'csv'
require 'benchmark'
# Axlsx::trust_input = true
row = []
input1 = (32..126).to_a.pack('U*').chars.to_a # these will need to be escaped
input2 = (65..122).to_a.pack('U*').chars.to_a # these do not need to be escaped
10.times { row << input1.shuffle.join }
10.times { row << input2.shuffle.join }
times = 3000

Benchmark.bmbm(30) do |x|
  x.report('axlsx_noautowidth') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_autowidth = false
    p.serialize("tmp/example_noautowidth.xlsx")
  end

  x.report('axlsx_autowidth') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("tmp/example_autowidth.xlsx")
  end

  x.report('axlsx_shared') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.use_shared_strings = true
    p.serialize("tmp/example_shared.xlsx")
  end

  x.report('axlsx_stream') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    s = p.to_stream
    File.binwrite('tmp/example_streamed.xlsx', s.read)
  end

  x.report('axlsx_zip_command') do
    p = Axlsx::Package.new
    p.workbook do |wb|
      wb.add_worksheet do |sheet|
        times.times do
          sheet << row
        end
      end
    end
    p.serialize("tmp/example_zip_command.xlsx", zip_command: 'zip')
  end

  x.report('csv') do
    CSV.open("example.csv", "wb") do |csv|
      times.times do
        csv << row
      end
    end
  end
end
File.delete("tmp/example.csv", "tmp/example_streamed.xlsx", "tmp/example_shared.xlsx", "tmp/example_autowidth.xlsx", "tmp/example_noautowidth.xlsx", "tmp/example_zip_command.xlsx")
