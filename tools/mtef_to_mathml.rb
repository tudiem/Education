# Dùng: ruby mtef_to_mathml.rb file1.bin file2.bin ...  → in JSON {tên_file: mathml | null}
# Cần gem: mathtype_to_mathml (+ pry, do gói thiếu dependency).
require 'json'
require 'mathtype_to_mathml'
out = {}
ARGV.each do |f|
  begin
    mml = MathTypeToMathML::Converter.new(f).convert
    out[File.basename(f)] = (mml.nil? || mml.strip.empty?) ? nil : mml
  rescue Exception => e
    out[File.basename(f)] = nil
  end
end
puts JSON.generate(out)
