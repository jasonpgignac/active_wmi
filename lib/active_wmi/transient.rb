require 'active_wmi/base'
module ActiveWmi
  class Transient < Base
    class << self
      def primary_key
        nil
      end
      def primary_key=(value)
        raise TransientIdError("Assigned a primary key to an ActiveWMI:Transient sub-class")
      end
      def element_path(id)
        raise TransientIdError("ActiveWMI:Transient sub-classes do not have a WMI Path")
      end
      # Check create
      def delete(id)
        throw TransientIdError("Cannot delete by id, as ActiveWmi:Transient objects have no id")
      end
      def exists?(id)
        false
      end
      private
        def find_every(options)
          return Array.new()
        end
        def find_single(options)
          throw ResourceNotFound
        end
    end
    def id
      return nil
    end
    def id=(value)
      raise TransientIdError("Assigned a value to id in an instance of an ActiveWMI:Transient sub-class")
    end
    def save
      attributes.each do |property, value|
        wmi_object.send(property + "=", value)
      end
    end
  end
  
  class TransientIdError < StandardError
  end

end
