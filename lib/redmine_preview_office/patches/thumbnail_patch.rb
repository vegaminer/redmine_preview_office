# encoding: utf-8
#
# Redmine plugin to preview a Microsoft Office attachment file
#
# Copyright © 2018 Stephan Wenzel <stephan.wenzel@drwpatent.de>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
#

module RedminePreviewOffice
  module Patches 
    module ThumbnailPatch
      def self.included(base)
        base.extend(ClassMethods)
        base.send(:include, InstanceMethods)
        base.class_eval do
          unloadable 	
          
          # for those, who read and analyze code: I haven't figured it out yet how to unset 
          # a constant and how to patch a function, which has been defined as self.function()
          # in a base.class_eval block
          #
		  @REDMINE_PREVIEW_OFFICE_CONVERT_BIN = ('soffice').freeze
	      @TMP_LOCK = Mutex.new

		  # Generates a thumbnail for the source image to target
          def self.generate_preview_office(source, target )
            begin
              @TMP_LOCK.lock
              self.do_convert( source, target )
            ensure
              @TMP_LOCK.unlock
            end
          end

		  def self.do_convert(source, target )
            return target if File.exists?(target)

            target_dir = File.dirname(target)
            logger.debug( 'TARGET_DIR: ' + target_dir.to_s )
            FileUtils.mkdir_p target_dir unless File.exists?(target_dir)
  
            target_lock = target + '.lock'
            # File.open( target_lock, "wb" ) { |inf|
            #   if inf.flock( File::LOCK_EX|File::LOCK_NB ) != 0
            #@TMP_LOCK.lock
            return target if File.exists?(target)

	        #if File.exists?( target_lock )
            #  @TMP_LOCK.unlock
            #  logger.debug( 'Target lock file exists, waiting: ' + target.to_s )
            #  1.upto(30) do |n|
            #    return target if File.exists?( target )
            #    logger.debug( 'Sleeping: ' + n.to_s + ' seconds' )
            #    sleep 1 # second
            #  end
            #  logger.debug( 'Waiting for target timeout. Exiting' )
            #  return nil  
            #else
			#  File.atomic_write( target_lock ) {|file| file.write( "locked" ) }
            #  @TMP_LOCK.unlock
            #end
 
              
	        puts "------------------- FGFGFGFGGFGFGFGFGFGFGFGFGFGFGFGFGFGFG ----------------------"
            logger.debug( "SRC: " + source.to_s )
            logger.debug( "DST: " + target.to_s )
                
            stdout, stderr, status = Open3.capture3( { 'PATH' => ENV[ 'PATH' ] }, @REDMINE_PREVIEW_OFFICE_CONVERT_BIN, '--convert-to', 'pdf', '--outdir', target_dir, source )
                    
            logger.debug( 'STDERR: ' )
            logger.debug( stderr )
            logger.debug( 'STDOUT: ' )
            logger.debug( stdout )
            logger.debug( 'STATUS: ' )
            logger.debug( status )
			     
            tmp_target = target_dir + '/' + File.basename(source, File.extname(source)) + ".pdf"
            unless File.exists?( tmp_target) # status.success?
              # File.delete( target_lock )
              cmd = "PATH='#{ENV[ 'PATH' ]}  #{shell_quote @REDMINE_PREVIEW_OFFICE_CONVERT_BIN} --convert-to pdf --outdir  #{shell_quote target_dir} #{shell_quote source}"
              puts "-------------------------------------------------------------------------------"
   		      logger.debug( "Creating preview with libreoffice failed (#{$?}):\nCommand: #{cmd}" )
             # File.delete( target_lock )
              #@TMP_LOCK.unlock
			  return nil
            end
            tmp_target = target_dir + '/' + File.basename(source, File.extname(source)) + ".pdf"
            logger.debug( 'RENAME: ' + tmp_target + ' => ' + target )
            File.rename( tmp_target, target )
            logger.debug( 'RENAME OK: ' + target_lock )  
            #File.delete( target_lock )
            #@TMP_LOCK.unlock
            #logger.debug( 'RENAME OK: ' + target )
		    return target
		  end #def 
		                     
		  def self.libreoffice_available?
			return @libreoffice_available if defined?(@libreoffice_available)
			@libreoffice_available = system("#{shell_quote @REDMINE_PREVIEW_OFFICE_CONVERT_BIN} --version") rescue false
			logger.warn("Libre Office (#{@REDMINE_PREVIEW_OFFICE_CONVERT_BIN}) not available") unless @libreoffice_available
			@libreoffice_available
		  end

        end #base
      end #self

      module InstanceMethods          		  
      end #module
      
      module ClassMethods
      end #module
      
    end #module
  end #module
end #module

unless Redmine::Thumbnail.included_modules.include?(RedminePreviewOffice::Patches::ThumbnailPatch)
    Redmine::Thumbnail.send(:include, RedminePreviewOffice::Patches::ThumbnailPatch)
end


