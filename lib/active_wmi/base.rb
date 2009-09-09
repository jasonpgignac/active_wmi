require 'active_wmi/connection'
require 'cgi'
require 'set'

module ActiveWmi
  # ActiveWmi::Base is the main class for mapping Windows WMI resources as
  # models in a Rails application.
  # == Mapping
  #
  # Active WMI objects represent your WMI resources as manipulatable Ruby 
  # objects.  To map resources to Ruby objects, Active WMI needs a few pieces 
  # of information:
  # 1) site - This is the address of the server to be contacted
  # 2) namespace - This is the namespace within the server where the objects
  #   are to be found
  # 2) user - A user account with rights to the necessary objects
  # 3) password - The password to authenticate the user account.
  # 4) element_name - The name of the object in WMI. This will default to the
  #   name of the class, but this is usually inappropriate as WMI tends to use
  #   names that are inconvenient.
  # 5) primary_key - the field or fields that act as a primary key for the  
  #   record. Support is built in for comopound keys, as many records in WMI
  #   demand them
  # Example:
  #   class Computer < ActiveWmi::Base
  #     self.site             = "smsserver.sample.com"
  #     self.namespace        = "root\\sms\\site_100"
  #     self.user             = "test_user"
  #     self.password         = "p@55w0rd"
  #     self.element_name     = "SMS_R_System"
  #     self.set_primary_key  = :resourceID
  #
  # Now the Computer class is mapped to WMI resources corresponding to the
  # computer records on an SMS Server sms.xy.com. 
  #
  # As mentioned, support for compound keys is baked in. For example:
  #
  #   self.element_name = "SMS_G_System_PC_BIOS"
  #   self.set_primary_keys([:GroupID,:ResourceID])
  #
  # Finally, because the column names can be somewhat unwieldly in a WMI 
  # record, support for aliasing columns is built in. For example, in the 
  # computer record above, one could add:
  #   alias_column :user, :lastlogonusername
  #   alias_column :addresses, :ipaddresses
  #   alias_column :domain, :resourcedomainorworkgroup
  #   alias_column :mac_addresses, :macaddresses
  #
  # You can now use Active Wmi's lifecycles methods to manipulate resources.
  #
  # == Lifecycle methods
  #
  # Active WMI exposes methods for creating, finding, updating, and deleting 
  # resources from WMI.
  #
  #   ryan = Person.new(:first => 'Ryan', :last => 'Daigle')
  #   ryan.save                # => true
  #   ryan.id                  # => 78747
  #   Person.exists?(ryan.id)  # => true
  #   ryan.exists?             # => true
  #
  #   ryan = Person.find(78747)
  #   # Resource holding our newly created Person object
  #   ryan = Person.find(:first, :conditions => {:first => 'Ryan'})
  #
  #   ryan.first = 'Rizzle'
  #   ryan.save                # => true
  #
  #   ryan.destroy             # => true
  #
  # As you can see, these are very similar to Active Record's lifecycle 
  # methods for database records.
  # You can read more about each of these methods in their respective 
  # documentation.
  #
  # == Validations
  #
  # There is, as of yet, no validation functionality for Active WMI
  #
  # == Errors & Validation
  #
  # While some errors have been defined, error handling in Active WMI is still
  # fairly primitive. You will need to refer to Microsoft's documentation in
  # order to translate WMI errors.  
  #
  # === Timeouts
  #
  # While the timeout variable has been put in place, it is not yet functional
  # in ActiveWMI
  
  
  class Base
    # The logger for diagnosing and tracing Active WMI calls.
    cattr_accessor :logger

    class << self
      def convert_to_windows_date_if_datetime(value)
        return value unless value.respond_to?('strftime')
        value.strftime("%Y%m%d%H%M%S.000000+***")
      end
      def convert_to_datetime_if_windows_date(value)
        return value unless value.is_a?(String) 
        return value unless value =~ /\A(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2}).(\d{4})/
        DateTime.civil($1.to_i,$2.to_i,$3.to_i,$4.to_i,$5.to_i,$6.to_i)
      end 
      # Gets the address of the WMI Server to connect for this class.  The 
      # site variable is required for Active Wmi's mapping to work. 
      def site
        # Not using superclass_delegating_reader because don't want subclasses 
        # to modify superclass instance
        #
        if defined?(@site)
          @site
        elsif superclass != Object && superclass.site
          superclass.site.dup.freeze
        end
      end
      # Sets the address (IP or DNS) of the WMI resources to map for this 
      # class to the value in the +site+ argument.
      # The site variable is required for Active WMI's mapping to work.
      def site=(site)
        @connection = nil
        if site.nil?
          @site = nil
        else
          @site = site
        end
      end
      # Gets the WMI namespace resources are to be sought in. The namespace 
      # variable is required for Active WMI mapping to work.
      def namespace
        # Not using superclass_delegating_reader because don't want subclasses to modify superclass instance
        #
        if defined?(@namespace)
          @namespace
        elsif superclass != Object && superclass.site
          superclass.site.dup.freeze
        end
      end
      # Sets the namespace of the WMI resources to map for this class to the 
      # value in the +namespace+ argument. The namespace variable is required 
      # for Active WMI's mapping to work.
      def namespace=(namespace)
        @connection = nil
        if namespace.nil?
          @namespace = nil
        else
          @namespace = namespace
        end
      end

      # Gets the user for OLE authentication.
      def user
        # Not using superclass_delegating_reader. See +site+ for explanation
        if defined?(@user)
          @user
        elsif superclass != Object && superclass.user
          superclass.user.dup.freeze
        end
      end

      # Sets the user for OLE authentication.
      def user=(user)
        @connection = nil
        @user = user
      end

      # Gets the password for OLE authentication.
      def password
        # Not using superclass_delegating_reader. See +site+ for explanation
        if defined?(@password)
          @password
        elsif superclass != Object && superclass.password
          superclass.password.dup.freeze
        end
      end

      # Sets the password for OLE authentication.
      def password=(password)
        @connection = nil
        @password = password
      end

      # Applies an alias name to a WMI column. The first paramater is the
      # desired name, and the second is the WMI column name.
      def alias_column(pretty, original)
        self.column_aliases[pretty] = original
      end
      # Gets column alias hash table
      def column_aliases
        @column_aliases ||= Hash.new
      end

      # An instance of ActiveWmi::Connection that is the base \connection to 
      # the remote service.
      # The +refresh+ parameter toggles whether or not the \connection is 
      # refreshed at every request
      # or not (defaults to <tt>false</tt>).
      # INCOMPLETE - Should require namespace in connection, and should reset 
      # connection object if variables change. Also, have not tested refresh 
      # variable
      def connection(refresh = false)
        if defined?(@connection) || superclass == Object
          @connection = Connection.new(site) if refresh || @connection.nil?
          @connection.namespace = namespace
          @connection.user = user if user
          @connection.password = password if password
          @connection
        else
          superclass.connection
        end
      end

      # element_name should be provided exactly as it appears in the WMI 
      # specification, though it is not case sensitive. For instance, 
      # sms_r_system for a computer resource object in SMS. This is the only 
      # name by which the object will be known to WMI, so this override allows 
      # you to use a more sane name for your object than WMI provides.
      attr_accessor_with_default(:element_name)    { 
        to_s.split("::").last.underscore 
      } #:nodoc:
      # primary_key should contain the name of the WMI field that acts as a 
      # primary key for the object. Certain objects have only a compound 
      # primary key, and this can be set by using an array of keys. You will 
      # almost always want to set this variable, since WMI classes use the 
      # format class_id instead of id.
      attr_accessor_with_default(:primary_key, 'id') #:nodoc:
      
      alias_method :set_element_name, :element_name=  #:nodoc:
      # Gets the element path for the given ID in +id+.
      #
      # ==== Examples
      #   Computer.element_path(74939)
      #   # => "SMS_R_System.ResourceID=74939"
      #   ComputerBios.element_path([40133,74939])
      #   # => "SMS_G_System_PC_BIOS.GroupID=40133,ResourceID=74939"
      #
      def element_path(id)
        if (self.primary_key.is_a?(Array))
          key_selects = Array.new
          primary_key.each do |key|
            f_id = id.shift
            f_id = f_id.is_a?(String) ? "\"" + f_id + "\"" : f_id.to_s
            key_selects << key.to_s + "=" + f_id
          end
          return self.element_name + "." + key_selects.join(",")
        else
          f_id = id.is_a?(String) ? "\"" + id + "\"" : id.to_s
          return self.element_name.to_s + "." + self.primary_key.to_s + "=" + f_id
        end
      end

      # Gets the collection path for the ActiveWmi resources.
      # ==== Examples
      #   Computer.collection_path
      #   # => "SMS_R_System"
      #
      def collection_path(query_options = nil)
        self.element_name
      end

      alias_method :set_primary_key, :primary_key=  #:nodoc:

      # Creates a new resource instance and makes a request to the remote 
      # service that it be saved, making it equivalent to the following 
      # simultaneous calls:
      #
      #   ryan = Person.new(:first => 'ryan')
      #   ryan.save
      #
      # Returns the newly created resource.  If a failure has occurred an
      # exception will be raised (see <tt>save</tt>). 
      # ==== Examples
      #   my_person = Person.create(:name => 'Jeremy', 
      #                             :email => 'myname@nospam.com', 
      #                             :enabled => true)
      # INCOMPLETE - function is not yet integrated with column aliases.
      def create(attributes = {})
        returning(self.new(attributes)) { |res| res.save }
      end

      # Core method for finding resources.  Used similarly to Active Record's 
      # +find+ method.
      #
      # ==== Arguments
      # The first argument (unless it is the id) is considered to be the scope 
      # of the query.  That is, how many resources are returned from the 
      # request.  It can be one of the following.
      #
      # * <tt>:first</tt> - Returns the first resource found.
      # * <tt>:last</tt> - Returns the last resource found.
      # * <tt>:all</tt> - Returns every resource that matches the request.
      #
      def find(*arguments)
        scope   = arguments.slice!(0)
        options = arguments.slice!(0) || {}

        case scope
          when :all   then find_every(options)
          when :first then find_every(options).first
          when :last  then find_every(options).last
          else             find_single(scope)
        end
      end

      # Deletes the resources with the ID in the +id+ parameter.
      # INCOMPLETE - Documentation incomplete
      # ==== Examples
      #   Event.delete(2) # sends DELETE /events/2
      #
      #   Event.create(:name => 'Free Concert', :location => 'Community Center')
      #   my_event = Event.find(:first) # let's assume this is event with ID 7
      #   Event.delete(my_event.id) # sends DELETE /events/7
      #
      #   # Let's assume a request to events/5/cancel.xml
      #   Event.delete(params[:id]) # sends DELETE /events/5
      def delete(id)
        self.find(id).delete_
      end

      # Asserts the existence of a resource, returning <tt>true</tt> if the 
      # resource is found.
      #
      # ==== Examples
      #   note = Note.create(:title => 'Hello, world.')
      #   note.id # => 1
      #   Note.exists?(1) # => true
      #
      #   Note.exists(1349) # => false
      def exists?(id)
        id && !find_single(id).nil?
      rescue ActiveWmi::ResourceNotFound
        false
      end
      
      private
        
        # Find every resource
        def find_every(options)
          query_options = options[:conditions]
          query_where_clause = where_clause_from(query_options)
          query = "SELECT * FROM " + self.element_name + ( query_where_clause ? " WHERE " + query_where_clause : "")
          instantiate_collection(connection.find(query) || [])
        end

        # Find a single resource from the primary key
        def find_single(scope)
          path = element_path(scope)
          self.new(connection.get(path))
        end
        
        # Create a collection or record from results
        def instantiate_collection(collection)
          collection.collect! { |record| instantiate_record(record) }
        end
        def instantiate_record(record)
          return new(connection.get(record.path_.path))
        end
        
        # Builds the query string for the request.
        def where_clause_from(options)
          return nil unless options
          query = String.new()
          query = sanitize_wql(options)
          puts "NOW, we'll execute this: #{query}"
          return query
        end
        # Accepts an array, hash, or string of sql conditions and sanitizes
        # them into a valid WQL fragment.
        #   ["name='%s' and group_id='%s'", "foo'bar", 4]  returns  "name='foo''bar' and group_id='4'"
        #   { :name => "foo'bar", :group_id => 4 }  returns "name='foo''bar' and group_id='4'"
        #   "name='foo''bar' and group_id='4'" returns "name='foo''bar' and group_id='4'"
        def sanitize_wql(condition)
          case condition
            when Array; sanitize_sql_array(condition)
            when Hash;  sanitize_sql_hash(condition)
            else        condition
          end
        end

        # Sanitizes a hash of attribute/value pairs into WQL conditions.
        #   { :name => "foo'bar", :group_id => 4 }
        #     # => "name='foo''bar' and group_id= 4"
        #   { :status => nil, :group_id => [1,2,3] }
        #     # => "status IS NULL and group_id IN (1,2,3)"
        #   { :age => 13..18 }
        #     # => "age BETWEEN 13 AND 18"
        def sanitize_sql_hash(attrs)
          conditions = attrs.map do |attr, value|
            true_attr = column_aliases.has_key?(attr) ? column_aliases[attr] : attr
            "#{true_attr} #{attribute_condition(value)}"
          end.join(' AND ')
          replace_bind_variables(conditions, attrs.values)
        end

        # Accepts an array of conditions.  The array has each value
        # sanitized and interpolated into the sql statement.
        #   ["name='%s' and group_id='%s'", "foo'bar", 4]  returns  "name='foo''bar' and group_id='4'"
        def sanitize_sql_array(ary)
          statement, *values = ary
          if values.first.is_a?(Hash) and statement =~ /:\w+/
            replace_named_bind_variables(statement, values.first)
          elsif statement.include?('?')
            replace_bind_variables(statement, values)
          else
            statement % values.collect { |value| quote_string(value) }
          end
        end

        def attribute_condition(argument)
          case argument
            when nil   then "IS ?"
    #         INCOMPLETE - Not functional with arrays or ranges
    #        when Array then "IN (?)"
    #        when Range then "BETWEEN ? AND ?"
            else            "= ?"
          end
        end

        def replace_bind_variables(statement, values) #:nodoc:
          raise_if_bind_arity_mismatch(statement, statement.count('?'), values.size)
          bound = values.dup
          statement.gsub('?') { quote_string(bound.shift) }
        end

        def raise_if_bind_arity_mismatch(statement, expected, provided)
          unless expected == provided
            raise PreparedStatementInvalid, "wrong number of bind variables (#{provided} for #{expected}) in: #{statement}"
          end
        end

        def quote_string(value)
          # records are quoted as their primary key
          case value
            when String, ActiveSupport::Multibyte::Chars
              value = value.to_s
              return "'#{value}'"
            when NilClass                 then "NULL"
            when TrueClass                then (TRUE)
            when FalseClass               then (FALSE)
            when Float, Fixnum, Bignum    then value.to_s
            # INCOMPLETE - (?) BigDecimals need to be output in a non-normalized form and quoted.
            # when BigDecimal               then value.to_s('F')
            else
              if value.acts_like?(:date) || value.acts_like?(:time)
                "'#{quoted_date(value)}'"
              else
                "#{quoted_string_prefix}'#{quote_string(value.to_yaml)}'"
              end
          end
        end

    end

    # INCOMPLETE - Do we need to write a cleaner accessor?
    attr_accessor :attributes #:nodoc:
    
    # INCOMPLETE - Does not work!
    def attributes
      unless(@attributes)
        @attributes = Hash.new()
        self.properties_.each do |property|
          @attributes[property.name] = property.value
        end
      end
      @attributes
    end
    
    # Constructor method for new resources; the optional +attributes+ 
    # parameter takes a hash of attributes for the new resource.
    #
    # ==== Examples
    #   my_course = Course.new
    #   my_course.name = "Western Civilization"
    #   my_course.lecturer = "Don Trotter"
    #   my_course.save
    #
    #   my_other_course = Course.new(:name => "Philosophy: Reason and Being", :lecturer => "Ralph Cling")
    #   my_other_course.save
    
    def initialize(*new_item)
      if new_item.length == 1 && new_item[0].is_a?(WIN32OLE)
        @wmi_object = new_item[0]
      else
        @wmi_object = connection.get(self.class.element_name).spawnInstance_()
        unless new_item.length == 0
          if new_item[0].is_a?(Hash)
            new_item[0].each do |attribute, value|
              self.send(attribute.to_s + "=", value)
            end
          else
            raise(ArgumentError)
          end
        end
      end
      self.attributes
      self
    end

    # Returns a \clone of the resource that hasn't been assigned an +id+ yet and
    # is treated as a \new resource.
    #
    #   ryan = Person.find(1)
    #   not_ryan = ryan.clone
    #   not_ryan.new?  # => true
    #
    # INCOMPLETE - What dowe do with records wtih subrecords, such as rules?
    def clone
      # Clone all attributes except the pk
      
      cloned = attributes.reject {|k,v| k.downcase == self.class.primary_key.downcase}.inject({}) do |attrs, (k, v)|
        case v
          when Fixnum     then attrs[k] = v 
          when NilClass   then attrs[k] = nil
          else            attrs[k] = v.clone
        end
        attrs
      end
      # Form the new resource - bypass initialize of resource with 'new' as that will call 'load' which
      # attempts to convert hashes into member objects and arrays into collections of objects.  We want
      # the raw objects to be cloned so we bypass load by directly setting the attributes hash.
      resource = self.class.new(cloned)
      resource
    end

    # A method to determine if the resource a \new object (i.e., it has not been saved via OLE yet).
    #
    # ==== Examples
    #   not_new = Computer.create(:brand => 'Apple', :make => 'MacBook', :vendor => 'MacMall')
    #   not_new.new? # => false
    #
    #   is_new = Computer.new(:brand => 'IBM', :make => 'Thinkpad', :vendor => 'IBM')
    #   is_new.new? # => true
    #
    #   is_new.save
    #   is_new.new? # => false
    #
    def new?
      id.nil?
    end

    # Gets the primary key attribute of the resource.
    def id
      if primary_key.is_a?(Array)
        id = Array.new
        primary_key.each do |key|
          sub_id = self.send(key)
          return nil if (id == "")
          id << sub_id
        end
      else
        id = self.send(self.class.primary_key)
        return nil if  (id == "")
      end
      return id
    end
    
    def primary_key
      self.class.primary_key
    end

    # INCOMPLETE - Haven't tested this at all, even a little bit
    # Allows Active WMI objects to be used as parameters in Action Pack URL generation.
    def to_param
      id && id.to_s
    end

    # Test for equality.  Resource are equal if and only if +other+ is the same object or
    # is an instance of the same class, is not <tt>new?</tt>, and has the same +id+.
    #
    # ==== Examples
    #   ryan = Person.create(:name => 'Ryan')
    #   jamie = Person.create(:name => 'Jamie')
    #
    #   ryan == jamie
    #   # => false (Different name attribute and id)
    #
    #   ryan_again = Person.new(:name => 'Ryan')
    #   ryan == ryan_again
    #   # => false (ryan_again is new?)
    #
    #   ryans_clone = Person.create(:name => 'Ryan')
    #   ryan == ryans_clone
    #   # => false (Different id attributes)
    #
    #   ryans_twin = Person.find(ryan.id)
    #   ryan == ryans_twin
    #   # => true
    #
    def ==(other)
      other.equal?(self) || (other.instance_of?(self.class) && !other.new? && other.id == id)
    end

    # Tests for equality (delegates to ==).
    def eql?(other)
      self == other
    end

    # Delegates to id in order to allow two resources of the same type and \id to work with something like:
    #   [Person.find(1), Person.find(2)] & [Person.find(1), Person.find(4)] # => [Person.find(1)]
    def hash
      id.hash
    end

    # Duplicate the current resource without saving it.
    #
    # ==== Examples
    #   my_invoice = Invoice.create(:customer => 'That Company')
    #   next_invoice = my_invoice.dup
    #   next_invoice.new? # => true
    #
    #   next_invoice.save
    #   next_invoice == my_invoice # => false (different id attributes)
    #
    #   my_invoice.customer   # => That Company
    #   next_invoice.customer # => That Company
    def dup
      returning self.class.new do |resource|
        resource.attributes     = @attributes
      end
    end

    # A method to save or update a resource.  It delegates to +create+ if a new object, 
    # +update+ if it is existing. 
    #
    # ==== Examples
    #   my_company = Company.new(:name => 'RoleModel Software', :owner => 'Ken Auer', :size => 2)
    #   my_company.new? # => true
    #   my_company.save
    #   my_company.new? # => false
    #   my_company.size = 10
    #   my_company.save
    
    def save
      new? ? create : update
    end

    # Deletes the resource from the remote service.
    #
    # ==== Examples
    #   my_id = 3
    #   my_person = Person.find(my_id)
    #   my_person.destroy
    #
    #   new_person = Person.create(:name => 'James')
    #   new_id = new_person.id # => 7
    #   new_person.destroy
    def destroy
      self.delete_
    end

    # Evaluates to <tt>true</tt> if this resource is not <tt>new?</tt> and is
    # found on the remote service.  Using this method, you can check for
    # resources that may have been deleted between the object's instantiation
    # and actions on it.
    #
    # ==== Examples
    #   Person.create(:name => 'Theodore Roosevelt')
    #   that_guy = Person.find(:first)
    #   that_guy.exists? # => true
    #
    #   that_lady = Person.new(:name => 'Paul Bean')
    #   that_lady.exists? # => false
    #
    #   guys_id = that_guy.id
    #   Person.delete(guys_id)
    #   that_guy.exists? # => false
    def exists?
      !new? && self.class.exists?(to_param, :params => prefix_options)
    end

    # A method to convert the the resource to an XML string.
    # INCOMPLETE
    # ==== Options
    # The +options+ parameter is handed off to the +to_xml+ method on each
    # attribute, so it has the same options as the +to_xml+ methods in
    # Active Support.
    #
    # * <tt>:indent</tt> - Set the indent level for the XML output (default is +2+).
    # * <tt>:dasherize</tt> - Boolean option to determine whether or not element names should
    #   replace underscores with dashes (default is <tt>false</tt>).
    # * <tt>:skip_instruct</tt> - Toggle skipping the +instruct!+ call on the XML builder
    #   that generates the XML declaration (default is <tt>false</tt>).
    #
    # ==== Examples
    #   my_group = SubsidiaryGroup.find(:first)
    #   my_group.to_xml
    #   # => <?xml version="1.0" encoding="UTF-8"?>
    #   #    <subsidiary_group> [...] </subsidiary_group>
    #
    #   my_group.to_xml(:dasherize => true)
    #   # => <?xml version="1.0" encoding="UTF-8"?>
    #   #    <subsidiary-group> [...] </subsidiary-group>
    #
    #   my_group.to_xml(:skip_instruct => true)
    #   # => <subsidiary_group> [...] </subsidiary_group>
    def to_xml(options={})
      attributes.to_xml({:root => self.class.element_name}.merge(options))
    end

    # Returns a JSON string representing the model. Some configuration is
    # available through +options+.
    # INCOMPLETE
    # ==== Options
    # The +options+ are passed to the +to_json+ method on each
    # attribute, so the same options as the +to_json+ methods in
    # Active Support.
    #
    # * <tt>:only</tt> - Only include the specified attribute or list of
    #   attributes in the serialized output. Attribute names must be specified
    #   as strings.
    # * <tt>:except</tt> - Do not include the specified attribute or list of
    #   attributes in the serialized output. Attribute names must be specified
    #   as strings.
    #
    # ==== Examples
    #   person = Person.new(:first_name => "Jim", :last_name => "Smith")
    #   person.to_json
    #   # => {"first_name": "Jim", "last_name": "Smith"}
    #
    #   person.to_json(:only => ["first_name"])
    #   # => {"first_name": "Jim"}
    #
    #   person.to_json(:except => ["first_name"])
    #   # => {"last_name": "Smith"}
    def to_json(options={})
      attributes.to_json(options)
    end

    # INCOPMLETE - Doesn't work
    # A method to \reload the attributes of this object from the remote web service.
    #
    # ==== Examples
    #   my_branch = Branch.find(:first)
    #   my_branch.name # => "Wislon Raod"
    #
    #   # Another client fixes the typo...
    #
    #   my_branch.name # => "Wislon Raod"
    #   my_branch.reload
    #   my_branch.name # => "Wilson Road"
    def reload
      @attributes = Hash.new()
      @wmi_object = self.class.find(self.to_param).wmi_object
      self.properties_.each do |property|
        @attributes[property.name] = property.value
      end
    end

    # INCOMPLETE = Doesn't work
    # A method to manually load attributes from a \hash. Recursively loads collections of
    # resources.  This method is called in +initialize+ and +create+ when a \hash of attributes
    # is provided.
    #
    # ==== Examples
    #   my_attrs = {:name => 'J&J Textiles', :industry => 'Cloth and textiles'}
    #   my_attrs = {:name => 'Marty', :colors => ["red", "green", "blue"]}
    #
    #   the_supplier = Supplier.find(:first)
    #   the_supplier.name # => 'J&M Textiles'
    #   the_supplier.load(my_attrs)
    #   the_supplier.name('J&J Textiles')
    #
    #   # These two calls are the same as Supplier.new(my_attrs)
    #   my_supplier = Supplier.new
    #   my_supplier.load(my_attrs)
    #
    #   # These three calls are the same as Supplier.create(my_attrs)
    #   your_supplier = Supplier.new
    #   your_supplier.load(my_attrs)
    #   your_supplier.save
    def load(object)
      if (object.is_a?(Hash))
        attributes = object
      else
        attributes = Hash.new()
        object.properties_.each do |property|
          attributes[property.name] = property.value
        end
      end
      attributes.each do |key, value|
        self.send
        @attributes[key.to_s] =
          case value
            when Array
              resource = find_or_create_resource_for_collection(key)
              value.map { |attrs| attrs.is_a?(String) ? attrs.dup : resource.new(attrs) }
            when Hash
              resource = find_or_create_resource_for(key)
              resource.new(value)
            else
              value.dup rescue value
          end
      end
      self
    end

    # For checking <tt>respond_to?</tt> without searching the attributes (which is faster).
    alias_method :respond_to_without_attributes?, :respond_to?

    # A method to determine if an object responds to a message (e.g., a method call). In Active Resource, a Person object with a
    # +name+ attribute can answer <tt>true</tt> to <tt>my_person.respond_to?(:name)</tt>, <tt>my_person.respond_to?(:name=)</tt>, and
    # <tt>my_person.respond_to?(:name?)</tt>.
    def respond_to?(method, include_priv = false)
      method_name = method.to_s
      if attributes.nil?
        return super
      elsif attributes.has_key?(method_name)
        return true
      elsif ['?','='].include?(method_name.last) && attributes.has_key?(method_name.first(-1))
        return true
      end
      # super must be called at the end of the method, because the inherited respond_to?
      # would return true for generated readers, even if the attribute wasn't present
      super
    end

    def wmi_object
      @wmi_object ||= self.connection.get_wmi_object(self)
    end
    protected
      def connection(refresh = false)
        self.class.connection(refresh)
      end

      # Update the resource on the remote service.
      def update
        wmi_object.put_
      end

      # Create the new resource.
      def create
        path = wmi_object.put_.path
        @wmi_object = (connection.get(path))        
      end

      def element_path(options = nil)
        self.class.element_path(to_param)
      end
      def collection_path(options = nil)
        self.class.collection_path(options)
      end

    private
      def column_aliases
        self.class.column_aliases
      end
      def method_missing(method_symbol, *arguments) #:nodoc:
        if (method_symbol.to_s =~ /(.*)=\Z/)
          method_symbol =$1.to_sym
          setter = true
        end
        if (column_aliases.key?(method_symbol))
          method_symbol = column_aliases[method_symbol]
        end
        method_name = method_symbol.to_s + (setter ? "=" : "")
        arguments.map! do |arg|
          self.class.convert_to_windows_date_if_datetime(arg)
        end
        response = wmi_object.send(method_name, *arguments)
        #INCOMPLETE - Should parse through structures, as well (Hashes, Arrays)
        self.class.convert_to_datetime_if_windows_date(response)
      end

  end
end
