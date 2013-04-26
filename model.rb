require 'ffi'
require 'singleton'

class ModelShim

  # WARNING: this is not thread safe
  def initialize
    reset
  end

  def reset
    Model.reset
  end

  def method_missing(name, *arguments)
    if arguments.size == 0
      get(name)
    elsif arguments.size == 1
      set(name, arguments.first)
    else
      super
    end 
  end

  def get(name)
    return 0 unless Model.respond_to?(name)
    ruby_value_from_excel_value(Model.send(name))
  end

  def ruby_value_from_excel_value(excel_value)
    case excel_value[:type]
    when :ExcelNumber; excel_value[:number]
    when :ExcelString; excel_value[:string].read_string.force_encoding("utf-8")
    when :ExcelBoolean; excel_value[:number] == 1
    when :ExcelEmpty; nil
    when :ExcelRange
      r = excel_value[:rows]
      c = excel_value[:columns]
      p = excel_value[:array]
      s = Model::ExcelValue.size
      a = Array.new(r) { Array.new(c) }
      (0...r).each do |row|
        (0...c).each do |column|
          a[row][column] = ruby_value_from_excel_value(Model::ExcelValue.new(p + (((row*c)+column)*s)))
        end
      end 
      return a
    when :ExcelError; [:value,:name,:div0,:ref,:na][excel_value[:number]]
    else
      raise Exception.new("ExcelValue type #{excel_value[:type].inspect} not recognised")
    end
  end

  def set(name, ruby_value)
    name = name.to_s
    name = "set_#{name[0..-2]}" if name.end_with?('=')
    return false unless Model.respond_to?(name)
    Model.send(name, excel_value_from_ruby_value(ruby_value))
  end

  def excel_value_from_ruby_value(ruby_value, excel_value = Model::ExcelValue.new)
    case ruby_value
    when Numeric
      excel_value[:type] = :ExcelNumber
      excel_value[:number] = ruby_value
    when String
      excel_value[:type] = :ExcelString
      excel_value[:string] = FFI::MemoryPointer.from_string(ruby_value.encode('utf-8'))
    when TrueClass, FalseClass
      excel_value[:type] = :ExcelBoolean
      excel_value[:number] = ruby_value ? 1 : 0
    when nil
      excel_value[:type] = :ExcelEmpty
    when Array
      excel_value[:type] = :ExcelRange
      # Presumed to be a row unless specified otherwise
      if ruby_value.first.is_a?(Array)
        excel_value[:rows] = ruby_value.size
        excel_value[:columns] = ruby_value.first.size
      else
        excel_value[:rows] = 1
        excel_value[:columns] = ruby_value.size
      end
      ruby_values = ruby_value.flatten
      pointer = FFI::MemoryPointer.new(Model::ExcelValue, ruby_values.size)
      excel_value[:array] = pointer
      ruby_values.each.with_index do |v,i|
        excel_value_from_ruby_value(v, Model::ExcelValue.new(pointer[i]))
      end
    when Symbol
      excel_value[:type] = :ExcelError
      excel_value[:number] = [:value, :name, :div0, :ref, :na].index(ruby_value)
    else
      raise Exception.new("Ruby value #{ruby_value.inspect} not translatable into excel")
    end
    excel_value
  end

end
    

module Model
  extend FFI::Library
  ffi_lib  File.join(File.dirname(__FILE__),FFI.map_library_name('model'))
  ExcelType = enum :ExcelEmpty, :ExcelNumber, :ExcelString, :ExcelBoolean, :ExcelError, :ExcelRange
                
  class ExcelValue < FFI::Struct
    layout :type, ExcelType,
  	       :number, :double,
  	       :string, :pointer,
         	 :array, :pointer,
           :rows, :int,
           :columns, :int             
  end
  

  # use this function to reset all cell values
  attach_function 'reset', [], :void

  # start of Model
  attach_function 'set_model_b9', [ExcelValue.by_value], :void
  attach_function 'set_model_b8', [ExcelValue.by_value], :void
  attach_function 'model_b13', [], ExcelValue.by_value
  attach_function 'model_b37', [], ExcelValue.by_value
  attach_function 'model_b48', [], ExcelValue.by_value
  attach_function 'model_c48', [], ExcelValue.by_value
  attach_function 'model_d48', [], ExcelValue.by_value
  attach_function 'model_e48', [], ExcelValue.by_value
  attach_function 'model_f48', [], ExcelValue.by_value
  attach_function 'model_g48', [], ExcelValue.by_value
  attach_function 'model_h48', [], ExcelValue.by_value
  attach_function 'model_i48', [], ExcelValue.by_value
  attach_function 'model_j48', [], ExcelValue.by_value
  attach_function 'model_k48', [], ExcelValue.by_value
  attach_function 'model_l48', [], ExcelValue.by_value
  attach_function 'model_m48', [], ExcelValue.by_value
  attach_function 'model_n48', [], ExcelValue.by_value
  attach_function 'model_o48', [], ExcelValue.by_value
  attach_function 'model_p48', [], ExcelValue.by_value
  attach_function 'model_q48', [], ExcelValue.by_value
  attach_function 'model_r48', [], ExcelValue.by_value
  attach_function 'model_s48', [], ExcelValue.by_value
  attach_function 'model_t48', [], ExcelValue.by_value
  attach_function 'model_u48', [], ExcelValue.by_value
  attach_function 'model_v48', [], ExcelValue.by_value
  attach_function 'model_w48', [], ExcelValue.by_value
  attach_function 'model_x48', [], ExcelValue.by_value
  attach_function 'model_y48', [], ExcelValue.by_value
  attach_function 'model_z48', [], ExcelValue.by_value
  attach_function 'model_aa48', [], ExcelValue.by_value
  attach_function 'model_ab48', [], ExcelValue.by_value
  attach_function 'model_ac48', [], ExcelValue.by_value
  attach_function 'model_ad48', [], ExcelValue.by_value
  attach_function 'model_ae48', [], ExcelValue.by_value
  attach_function 'model_af48', [], ExcelValue.by_value
  attach_function 'model_ag48', [], ExcelValue.by_value
  attach_function 'model_ah48', [], ExcelValue.by_value
  attach_function 'model_ai48', [], ExcelValue.by_value
  attach_function 'model_aj48', [], ExcelValue.by_value
  attach_function 'model_ak48', [], ExcelValue.by_value
  attach_function 'model_al48', [], ExcelValue.by_value
  attach_function 'model_am48', [], ExcelValue.by_value
  attach_function 'model_an48', [], ExcelValue.by_value
  attach_function 'model_b32', [], ExcelValue.by_value
  attach_function 'model_b31', [], ExcelValue.by_value
  attach_function 'model_b4', [], ExcelValue.by_value
  attach_function 'model_f3', [], ExcelValue.by_value
  attach_function 'model_b3', [], ExcelValue.by_value
  attach_function 'model_b53', [], ExcelValue.by_value
  attach_function 'model_c53', [], ExcelValue.by_value
  attach_function 'model_d53', [], ExcelValue.by_value
  attach_function 'model_e53', [], ExcelValue.by_value
  attach_function 'model_f53', [], ExcelValue.by_value
  attach_function 'model_g53', [], ExcelValue.by_value
  attach_function 'model_h53', [], ExcelValue.by_value
  attach_function 'model_i53', [], ExcelValue.by_value
  attach_function 'model_j53', [], ExcelValue.by_value
  attach_function 'model_k53', [], ExcelValue.by_value
  attach_function 'model_l53', [], ExcelValue.by_value
  attach_function 'model_m53', [], ExcelValue.by_value
  attach_function 'model_n53', [], ExcelValue.by_value
  attach_function 'model_o53', [], ExcelValue.by_value
  attach_function 'model_p53', [], ExcelValue.by_value
  attach_function 'model_q53', [], ExcelValue.by_value
  attach_function 'model_r53', [], ExcelValue.by_value
  attach_function 'model_s53', [], ExcelValue.by_value
  attach_function 'model_t53', [], ExcelValue.by_value
  attach_function 'model_u53', [], ExcelValue.by_value
  attach_function 'model_v53', [], ExcelValue.by_value
  attach_function 'model_w53', [], ExcelValue.by_value
  attach_function 'model_x53', [], ExcelValue.by_value
  attach_function 'model_y53', [], ExcelValue.by_value
  attach_function 'model_z53', [], ExcelValue.by_value
  attach_function 'model_aa53', [], ExcelValue.by_value
  attach_function 'model_ab53', [], ExcelValue.by_value
  attach_function 'model_ac53', [], ExcelValue.by_value
  attach_function 'model_ad53', [], ExcelValue.by_value
  attach_function 'model_ae53', [], ExcelValue.by_value
  attach_function 'model_af53', [], ExcelValue.by_value
  attach_function 'model_ag53', [], ExcelValue.by_value
  attach_function 'model_ah53', [], ExcelValue.by_value
  attach_function 'model_ai53', [], ExcelValue.by_value
  attach_function 'model_aj53', [], ExcelValue.by_value
  attach_function 'model_ak53', [], ExcelValue.by_value
  attach_function 'model_al53', [], ExcelValue.by_value
  attach_function 'model_am53', [], ExcelValue.by_value
  attach_function 'model_an53', [], ExcelValue.by_value
  attach_function 'model_b52', [], ExcelValue.by_value
  attach_function 'model_c52', [], ExcelValue.by_value
  attach_function 'model_d52', [], ExcelValue.by_value
  attach_function 'model_e52', [], ExcelValue.by_value
  attach_function 'model_f52', [], ExcelValue.by_value
  attach_function 'model_g52', [], ExcelValue.by_value
  attach_function 'model_h52', [], ExcelValue.by_value
  attach_function 'model_i52', [], ExcelValue.by_value
  attach_function 'model_j52', [], ExcelValue.by_value
  attach_function 'model_k52', [], ExcelValue.by_value
  attach_function 'model_l52', [], ExcelValue.by_value
  attach_function 'model_m52', [], ExcelValue.by_value
  attach_function 'model_n52', [], ExcelValue.by_value
  attach_function 'model_o52', [], ExcelValue.by_value
  attach_function 'model_p52', [], ExcelValue.by_value
  attach_function 'model_q52', [], ExcelValue.by_value
  attach_function 'model_r52', [], ExcelValue.by_value
  attach_function 'model_s52', [], ExcelValue.by_value
  attach_function 'model_t52', [], ExcelValue.by_value
  attach_function 'model_u52', [], ExcelValue.by_value
  attach_function 'model_v52', [], ExcelValue.by_value
  attach_function 'model_w52', [], ExcelValue.by_value
  attach_function 'model_x52', [], ExcelValue.by_value
  attach_function 'model_y52', [], ExcelValue.by_value
  attach_function 'model_z52', [], ExcelValue.by_value
  attach_function 'model_aa52', [], ExcelValue.by_value
  attach_function 'model_ab52', [], ExcelValue.by_value
  attach_function 'model_ac52', [], ExcelValue.by_value
  attach_function 'model_ad52', [], ExcelValue.by_value
  attach_function 'model_ae52', [], ExcelValue.by_value
  attach_function 'model_af52', [], ExcelValue.by_value
  attach_function 'model_ag52', [], ExcelValue.by_value
  attach_function 'model_ah52', [], ExcelValue.by_value
  attach_function 'model_ai52', [], ExcelValue.by_value
  attach_function 'model_aj52', [], ExcelValue.by_value
  attach_function 'model_ak52', [], ExcelValue.by_value
  attach_function 'model_al52', [], ExcelValue.by_value
  attach_function 'model_am52', [], ExcelValue.by_value
  attach_function 'model_an52', [], ExcelValue.by_value
  attach_function 'model_f6', [], ExcelValue.by_value
  attach_function 'model_f7', [], ExcelValue.by_value
  attach_function 'model_b50', [], ExcelValue.by_value
  attach_function 'model_c50', [], ExcelValue.by_value
  attach_function 'model_d50', [], ExcelValue.by_value
  attach_function 'model_e50', [], ExcelValue.by_value
  attach_function 'model_f50', [], ExcelValue.by_value
  attach_function 'model_g50', [], ExcelValue.by_value
  attach_function 'model_h50', [], ExcelValue.by_value
  attach_function 'model_i50', [], ExcelValue.by_value
  attach_function 'model_j50', [], ExcelValue.by_value
  attach_function 'model_k50', [], ExcelValue.by_value
  attach_function 'model_l50', [], ExcelValue.by_value
  attach_function 'model_m50', [], ExcelValue.by_value
  attach_function 'model_n50', [], ExcelValue.by_value
  attach_function 'model_o50', [], ExcelValue.by_value
  attach_function 'model_p50', [], ExcelValue.by_value
  attach_function 'model_q50', [], ExcelValue.by_value
  attach_function 'model_r50', [], ExcelValue.by_value
  attach_function 'model_s50', [], ExcelValue.by_value
  attach_function 'model_t50', [], ExcelValue.by_value
  attach_function 'model_u50', [], ExcelValue.by_value
  attach_function 'model_v50', [], ExcelValue.by_value
  attach_function 'model_w50', [], ExcelValue.by_value
  attach_function 'model_x50', [], ExcelValue.by_value
  attach_function 'model_y50', [], ExcelValue.by_value
  attach_function 'model_z50', [], ExcelValue.by_value
  attach_function 'model_aa50', [], ExcelValue.by_value
  attach_function 'model_ab50', [], ExcelValue.by_value
  attach_function 'model_ac50', [], ExcelValue.by_value
  attach_function 'model_ad50', [], ExcelValue.by_value
  attach_function 'model_ae50', [], ExcelValue.by_value
  attach_function 'model_af50', [], ExcelValue.by_value
  attach_function 'model_ag50', [], ExcelValue.by_value
  attach_function 'model_ah50', [], ExcelValue.by_value
  attach_function 'model_ai50', [], ExcelValue.by_value
  attach_function 'model_aj50', [], ExcelValue.by_value
  attach_function 'model_ak50', [], ExcelValue.by_value
  attach_function 'model_al50', [], ExcelValue.by_value
  attach_function 'model_am50', [], ExcelValue.by_value
  attach_function 'model_an50', [], ExcelValue.by_value
  attach_function 'model_b51', [], ExcelValue.by_value
  attach_function 'model_c51', [], ExcelValue.by_value
  attach_function 'model_d51', [], ExcelValue.by_value
  attach_function 'model_e51', [], ExcelValue.by_value
  attach_function 'model_f51', [], ExcelValue.by_value
  attach_function 'model_g51', [], ExcelValue.by_value
  attach_function 'model_h51', [], ExcelValue.by_value
  attach_function 'model_i51', [], ExcelValue.by_value
  attach_function 'model_j51', [], ExcelValue.by_value
  attach_function 'model_k51', [], ExcelValue.by_value
  attach_function 'model_l51', [], ExcelValue.by_value
  attach_function 'model_m51', [], ExcelValue.by_value
  attach_function 'model_n51', [], ExcelValue.by_value
  attach_function 'model_o51', [], ExcelValue.by_value
  attach_function 'model_p51', [], ExcelValue.by_value
  attach_function 'model_q51', [], ExcelValue.by_value
  attach_function 'model_r51', [], ExcelValue.by_value
  attach_function 'model_s51', [], ExcelValue.by_value
  attach_function 'model_t51', [], ExcelValue.by_value
  attach_function 'model_u51', [], ExcelValue.by_value
  attach_function 'model_v51', [], ExcelValue.by_value
  attach_function 'model_w51', [], ExcelValue.by_value
  attach_function 'model_x51', [], ExcelValue.by_value
  attach_function 'model_y51', [], ExcelValue.by_value
  attach_function 'model_z51', [], ExcelValue.by_value
  attach_function 'model_aa51', [], ExcelValue.by_value
  attach_function 'model_ab51', [], ExcelValue.by_value
  attach_function 'model_ac51', [], ExcelValue.by_value
  attach_function 'model_ad51', [], ExcelValue.by_value
  attach_function 'model_ae51', [], ExcelValue.by_value
  attach_function 'model_af51', [], ExcelValue.by_value
  attach_function 'model_ag51', [], ExcelValue.by_value
  attach_function 'model_ah51', [], ExcelValue.by_value
  attach_function 'model_ai51', [], ExcelValue.by_value
  attach_function 'model_aj51', [], ExcelValue.by_value
  attach_function 'model_ak51', [], ExcelValue.by_value
  attach_function 'model_al51', [], ExcelValue.by_value
  attach_function 'model_am51', [], ExcelValue.by_value
  attach_function 'model_an51', [], ExcelValue.by_value
  attach_function 'model_b40', [], ExcelValue.by_value
  attach_function 'model_c40', [], ExcelValue.by_value
  attach_function 'model_d40', [], ExcelValue.by_value
  attach_function 'model_b89', [], ExcelValue.by_value
  attach_function 'model_c89', [], ExcelValue.by_value
  attach_function 'model_d89', [], ExcelValue.by_value
  attach_function 'model_e89', [], ExcelValue.by_value
  attach_function 'model_f89', [], ExcelValue.by_value
  attach_function 'model_g89', [], ExcelValue.by_value
  attach_function 'model_h89', [], ExcelValue.by_value
  attach_function 'model_i89', [], ExcelValue.by_value
  attach_function 'model_j89', [], ExcelValue.by_value
  attach_function 'model_k89', [], ExcelValue.by_value
  attach_function 'model_l89', [], ExcelValue.by_value
  attach_function 'model_m89', [], ExcelValue.by_value
  attach_function 'model_n89', [], ExcelValue.by_value
  attach_function 'model_o89', [], ExcelValue.by_value
  attach_function 'model_p89', [], ExcelValue.by_value
  attach_function 'model_q89', [], ExcelValue.by_value
  attach_function 'model_r89', [], ExcelValue.by_value
  attach_function 'model_s89', [], ExcelValue.by_value
  attach_function 'model_t89', [], ExcelValue.by_value
  attach_function 'model_u89', [], ExcelValue.by_value
  attach_function 'model_v89', [], ExcelValue.by_value
  attach_function 'model_w89', [], ExcelValue.by_value
  attach_function 'model_x89', [], ExcelValue.by_value
  attach_function 'model_y89', [], ExcelValue.by_value
  attach_function 'model_z89', [], ExcelValue.by_value
  attach_function 'model_aa89', [], ExcelValue.by_value
  attach_function 'model_ab89', [], ExcelValue.by_value
  attach_function 'model_ac89', [], ExcelValue.by_value
  attach_function 'model_ad89', [], ExcelValue.by_value
  attach_function 'model_ae89', [], ExcelValue.by_value
  attach_function 'model_af89', [], ExcelValue.by_value
  attach_function 'model_ag89', [], ExcelValue.by_value
  attach_function 'model_ah89', [], ExcelValue.by_value
  attach_function 'model_ai89', [], ExcelValue.by_value
  attach_function 'model_aj89', [], ExcelValue.by_value
  attach_function 'model_ak89', [], ExcelValue.by_value
  attach_function 'model_al89', [], ExcelValue.by_value
  attach_function 'model_am89', [], ExcelValue.by_value
  attach_function 'model_an89', [], ExcelValue.by_value
  attach_function 'model_b85', [], ExcelValue.by_value
  attach_function 'model_c85', [], ExcelValue.by_value
  attach_function 'model_d85', [], ExcelValue.by_value
  attach_function 'model_e85', [], ExcelValue.by_value
  attach_function 'model_f85', [], ExcelValue.by_value
  attach_function 'model_g85', [], ExcelValue.by_value
  attach_function 'model_h85', [], ExcelValue.by_value
  attach_function 'model_i85', [], ExcelValue.by_value
  attach_function 'model_j85', [], ExcelValue.by_value
  attach_function 'model_k85', [], ExcelValue.by_value
  attach_function 'model_l85', [], ExcelValue.by_value
  attach_function 'model_m85', [], ExcelValue.by_value
  attach_function 'model_n85', [], ExcelValue.by_value
  attach_function 'model_o85', [], ExcelValue.by_value
  attach_function 'model_p85', [], ExcelValue.by_value
  attach_function 'model_q85', [], ExcelValue.by_value
  attach_function 'model_r85', [], ExcelValue.by_value
  attach_function 'model_s85', [], ExcelValue.by_value
  attach_function 'model_t85', [], ExcelValue.by_value
  attach_function 'model_u85', [], ExcelValue.by_value
  attach_function 'model_v85', [], ExcelValue.by_value
  attach_function 'model_w85', [], ExcelValue.by_value
  attach_function 'model_x85', [], ExcelValue.by_value
  attach_function 'model_y85', [], ExcelValue.by_value
  attach_function 'model_z85', [], ExcelValue.by_value
  attach_function 'model_aa85', [], ExcelValue.by_value
  attach_function 'model_ab85', [], ExcelValue.by_value
  attach_function 'model_ac85', [], ExcelValue.by_value
  attach_function 'model_ad85', [], ExcelValue.by_value
  attach_function 'model_ae85', [], ExcelValue.by_value
  attach_function 'model_af85', [], ExcelValue.by_value
  attach_function 'model_ag85', [], ExcelValue.by_value
  attach_function 'model_ah85', [], ExcelValue.by_value
  attach_function 'model_ai85', [], ExcelValue.by_value
  attach_function 'model_aj85', [], ExcelValue.by_value
  attach_function 'model_ak85', [], ExcelValue.by_value
  attach_function 'model_al85', [], ExcelValue.by_value
  attach_function 'model_am85', [], ExcelValue.by_value
  attach_function 'model_an85', [], ExcelValue.by_value
  attach_function 'model_b12', [], ExcelValue.by_value
  attach_function 'model_b9', [], ExcelValue.by_value
  attach_function 'model_b11', [], ExcelValue.by_value
  attach_function 'model_b45', [], ExcelValue.by_value
  attach_function 'model_c45', [], ExcelValue.by_value
  attach_function 'model_b10', [], ExcelValue.by_value
  attach_function 'model_b44', [], ExcelValue.by_value
  attach_function 'model_c44', [], ExcelValue.by_value
  attach_function 'model_b56', [], ExcelValue.by_value
  attach_function 'model_c56', [], ExcelValue.by_value
  attach_function 'model_d56', [], ExcelValue.by_value
  attach_function 'model_e56', [], ExcelValue.by_value
  attach_function 'model_f56', [], ExcelValue.by_value
  attach_function 'model_g56', [], ExcelValue.by_value
  attach_function 'model_h56', [], ExcelValue.by_value
  attach_function 'model_i56', [], ExcelValue.by_value
  attach_function 'model_j56', [], ExcelValue.by_value
  attach_function 'model_k56', [], ExcelValue.by_value
  attach_function 'model_l56', [], ExcelValue.by_value
  attach_function 'model_m56', [], ExcelValue.by_value
  attach_function 'model_n56', [], ExcelValue.by_value
  attach_function 'model_o56', [], ExcelValue.by_value
  attach_function 'model_p56', [], ExcelValue.by_value
  attach_function 'model_q56', [], ExcelValue.by_value
  attach_function 'model_r56', [], ExcelValue.by_value
  attach_function 'model_s56', [], ExcelValue.by_value
  attach_function 'model_t56', [], ExcelValue.by_value
  attach_function 'model_u56', [], ExcelValue.by_value
  attach_function 'model_v56', [], ExcelValue.by_value
  attach_function 'model_w56', [], ExcelValue.by_value
  attach_function 'model_x56', [], ExcelValue.by_value
  attach_function 'model_y56', [], ExcelValue.by_value
  attach_function 'model_z56', [], ExcelValue.by_value
  attach_function 'model_aa56', [], ExcelValue.by_value
  attach_function 'model_ab56', [], ExcelValue.by_value
  attach_function 'model_ac56', [], ExcelValue.by_value
  attach_function 'model_ad56', [], ExcelValue.by_value
  attach_function 'model_ae56', [], ExcelValue.by_value
  attach_function 'model_af56', [], ExcelValue.by_value
  attach_function 'model_ag56', [], ExcelValue.by_value
  attach_function 'model_ah56', [], ExcelValue.by_value
  attach_function 'model_ai56', [], ExcelValue.by_value
  attach_function 'model_aj56', [], ExcelValue.by_value
  attach_function 'model_ak56', [], ExcelValue.by_value
  attach_function 'model_al56', [], ExcelValue.by_value
  attach_function 'model_am56', [], ExcelValue.by_value
  attach_function 'model_an56', [], ExcelValue.by_value
  attach_function 'model_b36', [], ExcelValue.by_value
  attach_function 'model_b35', [], ExcelValue.by_value
  attach_function 'model_b7', [], ExcelValue.by_value
  attach_function 'model_b34', [], ExcelValue.by_value
  attach_function 'model_b8', [], ExcelValue.by_value
  attach_function 'model_b49', [], ExcelValue.by_value
  attach_function 'model_c49', [], ExcelValue.by_value
  attach_function 'model_d49', [], ExcelValue.by_value
  attach_function 'model_e49', [], ExcelValue.by_value
  attach_function 'model_f49', [], ExcelValue.by_value
  attach_function 'model_g49', [], ExcelValue.by_value
  attach_function 'model_h49', [], ExcelValue.by_value
  attach_function 'model_i49', [], ExcelValue.by_value
  attach_function 'model_j49', [], ExcelValue.by_value
  attach_function 'model_k49', [], ExcelValue.by_value
  attach_function 'model_l49', [], ExcelValue.by_value
  attach_function 'model_m49', [], ExcelValue.by_value
  attach_function 'model_n49', [], ExcelValue.by_value
  attach_function 'model_o49', [], ExcelValue.by_value
  attach_function 'model_p49', [], ExcelValue.by_value
  attach_function 'model_q49', [], ExcelValue.by_value
  attach_function 'model_r49', [], ExcelValue.by_value
  attach_function 'model_s49', [], ExcelValue.by_value
  attach_function 'model_t49', [], ExcelValue.by_value
  attach_function 'model_u49', [], ExcelValue.by_value
  attach_function 'model_v49', [], ExcelValue.by_value
  attach_function 'model_w49', [], ExcelValue.by_value
  attach_function 'model_x49', [], ExcelValue.by_value
  attach_function 'model_y49', [], ExcelValue.by_value
  attach_function 'model_z49', [], ExcelValue.by_value
  attach_function 'model_aa49', [], ExcelValue.by_value
  attach_function 'model_ab49', [], ExcelValue.by_value
  attach_function 'model_ac49', [], ExcelValue.by_value
  attach_function 'model_ad49', [], ExcelValue.by_value
  attach_function 'model_ae49', [], ExcelValue.by_value
  attach_function 'model_af49', [], ExcelValue.by_value
  attach_function 'model_ag49', [], ExcelValue.by_value
  attach_function 'model_ah49', [], ExcelValue.by_value
  attach_function 'model_ai49', [], ExcelValue.by_value
  attach_function 'model_aj49', [], ExcelValue.by_value
  attach_function 'model_ak49', [], ExcelValue.by_value
  attach_function 'model_al49', [], ExcelValue.by_value
  attach_function 'model_am49', [], ExcelValue.by_value
  attach_function 'model_an49', [], ExcelValue.by_value
  attach_function 'model_b55', [], ExcelValue.by_value
  attach_function 'model_c55', [], ExcelValue.by_value
  attach_function 'model_d55', [], ExcelValue.by_value
  attach_function 'model_e55', [], ExcelValue.by_value
  attach_function 'model_f55', [], ExcelValue.by_value
  attach_function 'model_g55', [], ExcelValue.by_value
  attach_function 'model_h55', [], ExcelValue.by_value
  attach_function 'model_i55', [], ExcelValue.by_value
  attach_function 'model_j55', [], ExcelValue.by_value
  attach_function 'model_k55', [], ExcelValue.by_value
  attach_function 'model_l55', [], ExcelValue.by_value
  attach_function 'model_m55', [], ExcelValue.by_value
  attach_function 'model_n55', [], ExcelValue.by_value
  attach_function 'model_o55', [], ExcelValue.by_value
  attach_function 'model_p55', [], ExcelValue.by_value
  attach_function 'model_q55', [], ExcelValue.by_value
  attach_function 'model_r55', [], ExcelValue.by_value
  attach_function 'model_s55', [], ExcelValue.by_value
  attach_function 'model_t55', [], ExcelValue.by_value
  attach_function 'model_u55', [], ExcelValue.by_value
  attach_function 'model_v55', [], ExcelValue.by_value
  attach_function 'model_w55', [], ExcelValue.by_value
  attach_function 'model_x55', [], ExcelValue.by_value
  attach_function 'model_y55', [], ExcelValue.by_value
  attach_function 'model_z55', [], ExcelValue.by_value
  attach_function 'model_aa55', [], ExcelValue.by_value
  attach_function 'model_ab55', [], ExcelValue.by_value
  attach_function 'model_ac55', [], ExcelValue.by_value
  attach_function 'model_ad55', [], ExcelValue.by_value
  attach_function 'model_ae55', [], ExcelValue.by_value
  attach_function 'model_af55', [], ExcelValue.by_value
  attach_function 'model_ag55', [], ExcelValue.by_value
  attach_function 'model_ah55', [], ExcelValue.by_value
  attach_function 'model_ai55', [], ExcelValue.by_value
  attach_function 'model_aj55', [], ExcelValue.by_value
  attach_function 'model_ak55', [], ExcelValue.by_value
  attach_function 'model_al55', [], ExcelValue.by_value
  attach_function 'model_am55', [], ExcelValue.by_value
  attach_function 'model_an55', [], ExcelValue.by_value
  attach_function 'model_b54', [], ExcelValue.by_value
  attach_function 'model_c54', [], ExcelValue.by_value
  attach_function 'model_d54', [], ExcelValue.by_value
  attach_function 'model_e54', [], ExcelValue.by_value
  attach_function 'model_f54', [], ExcelValue.by_value
  attach_function 'model_g54', [], ExcelValue.by_value
  attach_function 'model_h54', [], ExcelValue.by_value
  attach_function 'model_i54', [], ExcelValue.by_value
  attach_function 'model_j54', [], ExcelValue.by_value
  attach_function 'model_k54', [], ExcelValue.by_value
  attach_function 'model_l54', [], ExcelValue.by_value
  attach_function 'model_m54', [], ExcelValue.by_value
  attach_function 'model_n54', [], ExcelValue.by_value
  attach_function 'model_o54', [], ExcelValue.by_value
  attach_function 'model_p54', [], ExcelValue.by_value
  attach_function 'model_q54', [], ExcelValue.by_value
  attach_function 'model_r54', [], ExcelValue.by_value
  attach_function 'model_s54', [], ExcelValue.by_value
  attach_function 'model_t54', [], ExcelValue.by_value
  attach_function 'model_u54', [], ExcelValue.by_value
  attach_function 'model_v54', [], ExcelValue.by_value
  attach_function 'model_w54', [], ExcelValue.by_value
  attach_function 'model_x54', [], ExcelValue.by_value
  attach_function 'model_y54', [], ExcelValue.by_value
  attach_function 'model_z54', [], ExcelValue.by_value
  attach_function 'model_aa54', [], ExcelValue.by_value
  attach_function 'model_ab54', [], ExcelValue.by_value
  attach_function 'model_ac54', [], ExcelValue.by_value
  attach_function 'model_ad54', [], ExcelValue.by_value
  attach_function 'model_ae54', [], ExcelValue.by_value
  attach_function 'model_af54', [], ExcelValue.by_value
  attach_function 'model_ag54', [], ExcelValue.by_value
  attach_function 'model_ah54', [], ExcelValue.by_value
  attach_function 'model_ai54', [], ExcelValue.by_value
  attach_function 'model_aj54', [], ExcelValue.by_value
  attach_function 'model_ak54', [], ExcelValue.by_value
  attach_function 'model_al54', [], ExcelValue.by_value
  attach_function 'model_am54', [], ExcelValue.by_value
  attach_function 'model_an54', [], ExcelValue.by_value
  # end of Model
  # Start of named references
  attach_function 'average_life_of_low_carbon_generation', [], ExcelValue.by_value
  attach_function 'ccs_by_2020', [], ExcelValue.by_value
  attach_function 'demand', [], ExcelValue.by_value
  attach_function 'electricity_demand_growth_rate', [], ExcelValue.by_value
  attach_function 'electricity_demand_in_2012', [], ExcelValue.by_value
  attach_function 'electricity_demand_in_2050', [], ExcelValue.by_value
  attach_function 'electricity_emissions_during_cb4', [], ExcelValue.by_value
  attach_function 'electrification_start_year', [], ExcelValue.by_value
  attach_function 'emissions', [], ExcelValue.by_value
  attach_function 'emissions_factor', [], ExcelValue.by_value
  attach_function 'emissions_factor_2030', [], ExcelValue.by_value
  attach_function 'emissions_factor_2050', [], ExcelValue.by_value
  attach_function 'high_carbon', [], ExcelValue.by_value
  attach_function 'high_carbon_ef', [], ExcelValue.by_value
  attach_function 'high_carbon_emissions_factor_2012', [], ExcelValue.by_value
  attach_function 'high_carbon_emissions_factor_2020', [], ExcelValue.by_value
  attach_function 'high_carbon_emissions_factor_2050', [], ExcelValue.by_value
  attach_function 'high_carbon_load_factor', [], ExcelValue.by_value
  attach_function 'low_carbon_load_factor', [], ExcelValue.by_value
  attach_function 'maximum_low_c', [], ExcelValue.by_value
  attach_function 'maximum_low_carbon_build_rate', [], ExcelValue.by_value
  attach_function 'maximum_low_carbon_build_rate_expansion', [], ExcelValue.by_value
  attach_function 'maxmean2012', [], ExcelValue.by_value
  attach_function 'maxmean2050', [], ExcelValue.by_value
  attach_function 'minimum_low_carbon_build_rate', [], ExcelValue.by_value
  attach_function 'minmean2012', [], ExcelValue.by_value
  attach_function 'minmean2050', [], ExcelValue.by_value
  attach_function 'net_increase_in_zero_carbon', [], ExcelValue.by_value
  attach_function 'nuclear_change_2012_2020', [], ExcelValue.by_value
  attach_function 'nuclear_in_2012', [], ExcelValue.by_value
  attach_function 'renewable_electricity_in_2020', [], ExcelValue.by_value
  attach_function 'renewables_in_2012', [], ExcelValue.by_value
  attach_function 'year_second_wave_of_building_starts', [], ExcelValue.by_value
  attach_function 'zero_carbon', [], ExcelValue.by_value
  attach_function 'zero_carbon_built', [], ExcelValue.by_value
  attach_function 'zero_carbon_decomissioned', [], ExcelValue.by_value
  attach_function 'set_maximum_low_carbon_build_rate', [ExcelValue.by_value], :void
  attach_function 'set_year_second_wave_of_building_starts', [ExcelValue.by_value], :void
  # End of named references
end
