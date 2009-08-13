require 'date'
require 'time'
require 'benchmark'
require 'timeout'
module ActiveWmi
  OLE_NOT_FOUND_ERROR_CODE = "80041002"
  
  class ConnectionError < StandardError # :nodoc:
    attr_reader :response

    def initialize(response, message = nil)
      @response = response
      @message  = message
    end

    def to_s
      "Failed with " + response.to_s + " : " + message.to_s
    end
  end

  # Raised when a Timeout::Error occurs.
  class TimeoutError < ConnectionError
    def initialize(message)
      @message = message
    end
    def to_s; @message ;end
  end

  # 3xx Redirection
  class Redirection < ConnectionError # :nodoc:
    def to_s; response['Location'] ? "#{super} => #{response['Location']}" : super; end
  end

  # 4xx Client Error
  class ClientError < ConnectionError; end # :nodoc:

  # 400 Bad Request
  class BadRequest < ClientError; end # :nodoc

  # 401 Unauthorized
  class UnauthorizedAccess < ClientError; end # :nodoc

  # 403 Forbidden
  class ForbiddenAccess < ClientError; end # :nodoc

  # 404 Not Found
  class ResourceNotFound < ClientError; end # :nodoc:

  # 409 Conflict
  class ResourceConflict < ClientError; end # :nodoc:

  # 5xx Server Error
  class ServerError < ConnectionError; end # :nodoc:

  # 405 Method Not Allowed
  class MethodNotAllowed < ClientError # :nodoc:
    def allowed_methods
      @response['Allow'].split(',').map { |verb| verb.strip.downcase.to_sym }
    end
  end

  # Class to handle connections to remote web services.
  # This class is used by ActiveResource::Base to interface with REST
  # services.
  class Connection
    attr_reader :site, :user, :password, :namespace
    
    # The +site+ parameter is required and will set the +site+
    # attribute to the address for the remote resource service.
    def initialize(site)
      raise ArgumentError, 'Missing site address' unless site
      @user = @password = nil
      self.site = site
    end

    # Set address for remote service.
    def site=(site)
      @site = site
    end
    
    def namespace=(namespace)
      @namespace = namespace
    end

    # Set user for remote service.
    def user=(user)
      @user = user
    end

    # Set password for remote service.
    def password=(password)
      @password = password
    end

    # Creates a WIN32OLE connection to a a subobject
    def get_wmi_object(object)
      @wmi_connection.get(object.class.element_name + "." + object.class.primary_key + "=\"" + object.id + "\"")
    end

    def find(query)
      wbemCollectedResponses = wmi.execQuery(query, "WQL", 0)
      responses = Array.new()
      wbemCollectedResponses.each do |response|
        responses << response
      end
      return responses
    end
    
    def get(item)
      wmi.get(item.to_s)
    rescue
      /OLE error code:(\d*) in (\w*)\n\s*(.*)\n/ =~ $!.to_s
      error_code = $1
      error_source = $2
      error_message = $3
      full_message = $!.to_s
      if(error_code == OLE_NOT_FOUND_ERROR_CODE)
        raise ResourceNotFound.new(error_code, error_message) 
      else
        raise
      end
    end
    
    private
      # INCOMPLETE
      # Makes request to remote service.
      def request(method, *arguments)
        logger.info "#{method.to_s.upcase} #{site.to_s}" if logger
        result = nil
        time = Benchmark.realtime { result = wmi.send(method, *arguments) }
        logger.info "--> %d %s (%d %.2fs)" % [result.code, result.message, result.body ? result.body.length : 0, time] if logger
        # handle_response(result)
      rescue Timeout::Error => e
        raise TimeoutError.new(e.message)
      end

      # Handles response and error codes from remote service.
      # INCOMPLETE
      def handle_response(response)
        case response.code.to_i
          when 301,302
            raise(Redirection.new(response))
          when 200...400
            response
          when 400
            raise(BadRequest.new(response))
          when 401
            raise(UnauthorizedAccess.new(response))
          when 403
            raise(ForbiddenAccess.new(response))
          when 404
            raise(ResourceNotFound.new(response))
          when 405
            raise(MethodNotAllowed.new(response))
          when 409
            raise(ResourceConflict.new(response))
          when 422
            raise(ResourceInvalid.new(response))
          when 401...500
            raise(ClientError.new(response))
          when 500...600
            raise(ServerError.new(response))
          else
            raise(ConnectionError.new(response, "Unknown response code: #{response.code}"))
        end
      end

      # Creates new WIN32OLE instance for communication with
      # remote service and resources.
      def wmi
        unless @wmi_connection
          wmi = WIN32OLE.new('WbemScripting.SWbemLocator')
          @wmi_connection = wmi.connectServer(@site.to_s, @namespace.to_s, @user, @password)
        end  
        return @wmi_connection
      end

      def logger #:nodoc:
        Base.logger
      end
  end
end
