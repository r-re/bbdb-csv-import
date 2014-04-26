;;; bbdb-csv-import.el --- import csv to bbdb version 3+

;; Copyright (C) 2014 by Ian Kelling

;; Maintainer: Ian Kelling <ian@iankelling.org>
;; Author: Ian Kelling <ian@iankelling.org>
;; Created: 1 Apr 2014
;; Version: 1.1
;; Package-Requires: ((pcsv "1.3.3") (dash "2.5.0") (bbdb "20140412.1949"))
;; Keywords: csv, util, bbdb
;; Homepage: https://gitlab.com/iankelling/bbdb-csv-import

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; Importer of csv (comma separated value) text into Emacs’s bbdb database,
;; version 3+. Works out of the box with csv exported from Thunderbird, Gmail,
;; Linkedin, Outlook.com/hotmail, and probably others. 
;; Easily extensible to handle new formats. 

;;; Installation:
;;
;; If you installed this file with a package manager, just
;; 
;; (require 'bbdb-csv-import)
;;
;; Else, note the min versions of dependencies above in "Package-Requires:",
;; and load this file. The exact minimum bbdb version is unknown, something 3+.
;;
;;; Basic Usage:
;;
;; Back up bbdb by copying `bbdb-file' in case things go wrong.
;;
;; Simply M-x `bbdb-csv-import-buffer' or `bbdb-csv-import-file'.
;; When called interactively, they prompt for file or buffer arguments.

;;; Advanced usage / notes:
;; Tested to work with thunderbird, gmail, linkedin, outlook.com/hotmail.com. For
;; those programs, if it's exporter has an option of what kind of csv format,
;; choose it's own native format if available, if not, choose an outlook
;; compatible format. If you're exporting from some other program and its csv
;; exporter claims outlook compatibility, there is a good chance it will work
;; out of the box.
;;
;; Duplicate contacts (according to email address) are skipped if
;; bbdb-allow-duplicates is nil (default). Any duplicates found are echoed at
;; the end of the import.
;;
;; If things don't work, you can probably fix it with a custom field mapping
;; variable. It should not be too hard. Use the existing tables as an
;; example. By default, we use a combination of most predefined mappings, and
;; look for all of their fields, but it is probably best to avoid that kind of
;; table when setting up your own as it is an unnecessary complexity in that
;; case. If you have a problem with data from a supported export program, start
;; by testing its specific mapping table instead of the combined one. Here is a
;; handy template to set each of the predefined mapping tables if you would
;; rather avoid the configure interface:
;; 
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-combined)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-thunderbird)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-gmail)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-gmail-typed-email)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-linkedin)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-outlook-web)
;; (setq bbdb-csv-import-mapping-table bbdb-csv-import-outlook-typed-email)
;; 
;; In addition to the examples, the doc string for `bbdb-create-internal' may
;; also be useful. Please send any new mapping tables to the maintainer listed
;; in this file. The maintainer should be able to help with any issues and may
;; create a new mapping table given sample data.
;;
;; Misc tips/troubleshooting:
;; - ASynK looks promising for syncing bbdb/google/outlook.
;; - The git repo contains a test folder with exactly tested version info and working
;;   test data.
;; - bbdb doesn't work if you delete the bbdb database file in
;;   the middle of an emacs session. If you want to empty the current bbdb database,
;;   do M-x bbdb then .* then C-u * d on the beginning of a record.
;; - After changing a mapping table variable, don't forget to re-execute
;;   (setq bbdb-csv-import-mapping-table ...) so that it propagates.


;;; Code:
(require 'pcsv)
(require 'dash)
(require 'bbdb-com)
(eval-when-compile (require 'cl))


(defconst bbdb-csv-import-thunderbird
  '((:namelist "First Name" "Last Name")
    (:name "Display Name")
    (:aka "Nickname")
    (:mail "Primary Email" "Secondary Email")
    (:phone "Work Phone" "Home Phone" "Fax Number" "Pager Number" "Mobile Number")
    (:address
     (("home address"
       (("Home Address" "Home Address 2")
        "Home City" "Home State"
        "Home ZipCode" "Home Country"))
      ("work address"
       (("Work Address" "Work Address 2")
        "Work City" "Work State"
        "Work ZipCode" "Work Country"))))
    (:organization "Organization")
    (:xfields "Web Page 1" "Web Page 2" "Birth Year" "Birth Month"
              "Birth Day" "Department" "Custom 1" "Custom 2" "Custom 3"
              "Custom 4" "Notes" "Job Title"))
  "Thunderbird csv format")

(defconst bbdb-csv-import-linkedin
  '((:namelist "First Name" "Middle Name" "Last Name")
    (:affix "Suffix")
    (:mail "E-mail Address" "E-mail 2 Address" "E-mail 3 Address")
    (:phone
     "Assistant's Phone" "Business Fax" "Business Phone"
     "Business Phone 2" "Callback" "Car Phone"
     "Company Main Phone" "Home Fax" "Home Phone"
     "Home Phone 2" "ISDN" "Mobile Phone"
     "Other Fax" "Other Phone" "Pager"
     "Primary Phone" "Radio Phone" "TTY/TDD Phone" "Telex")
    (:address
     (("business address"
       (("Business Street" "Business Street 2" "Business Street 3")
        "Business City" "Business State"
        "Business Postal Code" "Business Country"))
      ("home address"
       (("Home Street" "Home Street 2" "Home Street 3")
        "Home City" "Home State"
        "Home Postal Code" "Home Country"))
      ("other address"
       (("Other Street" "Other Street 2" "Other Street 3")
        "Other City" "Other State"
        "Other Postal Code" "Other Country"))))
    (:organization "Company")
    (:xfields
     "Department" "Job Title" "Assistant's Name"
     "Birthday" "Manager's Name" "Notes" "Other Address PO Box"
     "Spouse" "Web Page" "Personal Web Page"))
  "Linkedin export in the Outlook csv format.")


(defconst bbdb-csv-import-gmail
  '((:namelist "Given Name" "Family Name")
    (:name "Name")
    (:affix "Name Prefix" "Name Suffix")
    (:aka "Nickname")
    (:mail (repeat "E-mail 1 - Value"))
    (:phone (repeat ("Phone 1 - Type" "Phone 1 - Value")))
    (:address
     (repeat (("Address 1 - Type")
              (("Address 1 - Street" "Address 1 - PO Box" "Address 1 - Extended Address")
               "Address 1 - City" "Address 1 - Region"
               "Address 1 - Postal Code" "Address 1 - Country"))))
    (:organization (repeat "Organization 1 - Name"))
    (:xfields
     "Additional Name" "Yomi Name" "Given Name Yomi"
     "Additional Name Yomi" "Family Name Yomi" 
     "Initials" "Short Name" "Maiden Name" "Birthday"
     "Gender" "Location" "Billing Information"
     "Directory Server" "Mileage" "Occupation"
     "Hobby" "Sensitivity" "Priority"
     "Subject" "Notes" "Group Membership"
     ;; Gmail wouldn't let me add more than 1 organization in its web interface,
     ;; but no harm in looking for multiple since the field name implies the
     ;; possibility.
     (repeat  
      "Organization 1 - Type" "Organization 1 - Yomi Name"
      "Organization 1 - Title" "Organization 1 - Department"
      "Organization 1 - Symbol" "Organization 1 - Location"
      "Organization 1 - Job Description")
     (repeat ("Relation 1 - Type" "Relation 1 - Value"))
     (repeat ("Website 1 - Type" "Website 1 - Value"))
     (repeat ("Event 1 - Type" "Event 1 - Value"))
     (repeat ("Custom Field 1 - Type" "Custom Field 1 - Value"))))
  "Gmail csv export format. Note some fields don't map perfectly,
feel free to modify them as you wish. \"PO Box\" and \"Extended
Address\" are added as additional address street lines if they
exist. Some special name fields are made custom instead of put in
name, which gets a single string. We map Gmail's \"Name Prefix\"
and \"Name Suffix\" to bbdb's affix (a list of strings). We lose
the prefix/suffix label, but those are usually obvious.")


(defconst bbdb-csv-import-gmail-typed-email
  (append  (car (last bbdb-csv-import-gmail)) '((repeat "E-mail 1 - Type")))
  "Like the first Gmail mapping, but use custom fields to store
   Gmail's email labels. This is separate because I assume most
   people don't use those labels and using the default labels
   would create useless custom fields.")

(defconst bbdb-csv-import-outlook-web
  '((:namelist "First Name" "Middle Name" "Last Name")
    (:mail "E-mail Address" "E-mail 2 Address" "E-mail 3 Address")
    (:affix "Suffix")
    (:phone
     "Assistant's Phone" "Business Fax" "Business Phone"
     "Business Phone 2" "Callback" "Car Phone"
     "Company Main Phone" "Home Fax" "Home Phone"
     "Home Phone 2" "ISDN" "Mobile Phone"
     "Other Fax" "Other Phone" "Pager"
     "Primary Phone" "Radio Phone" "TTY/TDD Phone" "Telex")
    (:address
     (("business address"
       (("Business Street")
        "Business City" "Business State"
        "Business Postal Code" "Business Country"))
      ("home address"
       (("Home Street")
        "Home City" "Home State"
        "Home Postal Code" "Home Country"))
      ("other address"
       (("Other Street")
        "Other City" "Other State"
        "Other Postal Code" "Other Country"))))
    (:organization "Company")
    (:xfields
     "Anniversary" "Family Name Yomi" "Given Name Yomi"
     "Department" "Job Title" "Birthday" "Manager's Name" "Notes"
     "Spouse" "Web Page"))
  "Hotmail.com, outlook.com, live.com, etc.
Based on 'Export for outlook.com and other services',
not the export for Outlook 2010 and 2013.")

(defconst bbdb-csv-import-outlook-typed-email
  (append  (car (last bbdb-csv-import-outlook-web)) '((repeat "E-mail 1 - Type")))
  "Like bbdb-csv-import-gmail-typed-email, but for outlook-web.
Adds email labels as custom fields.")


(defun bbdb-csv-import-flatten1 (list)
  "Flatten LIST by 1 level."
  (--reduce-from (if (consp it)
                     (-concat acc it)
                   (-snoc acc it))
                 nil list))


(defun bbdb-csv-import-merge-map (root)
  "Combine two root mappings for making a combined mapping."
  (bbdb-csv-import-flatten1
   (list root
         (-distinct
          (append
           (cdr (assoc root bbdb-csv-import-thunderbird))
           (cdr (assoc root bbdb-csv-import-linkedin))
           (cdr (assoc root bbdb-csv-import-gmail))
           (cdr (assoc root bbdb-csv-import-outlook-web)))))))


(defconst bbdb-csv-import-combined
  (list
   ;; manually combined for proper ordering
   '(:namelist "First Name" "Given Name" "Middle Name" "Last Name" "Family Name")
   (bbdb-csv-import-merge-map :name)
   (bbdb-csv-import-merge-map :affix)
   (bbdb-csv-import-merge-map :aka)
   (bbdb-csv-import-merge-map :mail)
   (bbdb-csv-import-merge-map :phone)
   ;; manually combined the addresses. Because it was easier.
   '(:address
     (repeat (("Address 1 - Type")
              (("Address 1 - Street" "Address 1 - PO Box" "Address 1 - Extended Address")
               "Address 1 - City" "Address 1 - Region"
               "Address 1 - Postal Code" "Address 1 - Country")))
     (("business address"
       (("Business Street" "Business Street 2" "Business Street 3")
        "Business City" "Business State"
        "Business Postal Code" "Business Country"))
      ("home address"
       (("Home Street" "Home Street 2" "Home Street 3"
         "Home Address" "Home Address 2")
        "Home City" "Home State"
        "Home Postal Code" "Home ZipCode" "Home Country"))
      ("work address"
       (("Work Address" "Work Address 2")
        "Work City" "Work State"
        "Work ZipCode" "Work Country"))
      ("other address"
       (("Other Street" "Other Street 2" "Other Street 3")
        "Other City" "Other State"
        "Other Postal Code" "Other Country"))))
   (bbdb-csv-import-merge-map :organization)
   (bbdb-csv-import-merge-map :xfields)))

(defcustom bbdb-csv-import-mapping-table bbdb-csv-import-combined
  "The table which maps bbdb fields to csv fields. The default should work for most cases.
See the commentary section of this file for more details."
  :group 'bbdb-csv-import
  :type 'symbol)


(defun  bbdb-csv-import-expand-repeats (csv-fields list)
  "Return new list where elements from LIST in form (repeat elem1 ...)
become ((elem1 ...) [(elem2 ...)] ...) for as many repeating
numbered fields exist in the csv fields. elem can be a string or
a tree (a list with lists inside it)"
  (cl-flet ((replace-num (num string)
                         ;; in STRING, replace all groups of numbers with NUM
                         (replace-regexp-in-string "[0-9]+"
                                                   (number-to-string num)
                                                   string)))
    (--reduce-from
     (if (not (and (consp it) (eq (car it) 'repeat)))
         (cons it acc)
       (setq it (cdr it))
       (let* ((i 1)
              (first-field (car (flatten it))))
         (setq acc (cons it acc))
         ;; use first-field to test if there is another repetition.
         (while (member
                 (replace-num (setq i (1+ i)) first-field)
                 csv-fields)
           (cl-labels ((fun (cell)
                            (if (consp cell)
                                (mapcar #'fun cell)
                              (replace-num i cell))))
             (setq acc (cons (fun it) acc))))
         acc))
     nil list)))

(defun bbdb-csv-import-map-bbdb (csv-fields root)
  "ROOT is a root element from bbdb-csv-import-mapping-table. Get
the csv-fields for root in the mapping format, including variably
repeated ones. Flatten by one because repeated fields are put in
sub-lists, but after expanding them, that extra depth is no
longer useful. Small trade off: address mappings without 'repeat need
to be grouped in a list because they contain sublists that we
don't want flattened."
  (bbdb-csv-import-flatten1
   (bbdb-csv-import-expand-repeats
    csv-fields
    (cdr (assoc root bbdb-csv-import-mapping-table)))))

;;;###autoload
(defun bbdb-csv-import-file (filename)
  "Parse and import csv file FILENAME to bbdb."
  (interactive "fCSV file containg contact data: ")
  (bbdb-csv-import-buffer (find-file-noselect filename)))

;;;###autoload
(defun bbdb-csv-import-buffer (&optional buffer-or-name) 
  "Parse and import csv BUFFER-OR-NAME to bbdb.
Argument is a buffer or name of a buffer.
Defaults to current buffer."
  (interactive "bBuffer containing CSV contact data: ")
  (when (null bbdb-csv-import-mapping-table)
    (error "error: `bbdb-csv-import-mapping-table' is nil. Please set it and rerun."))
  (let* ((csv-data (pcsv-parse-buffer (get-buffer (or buffer-or-name (current-buffer)))))
         (csv-fields (car csv-data))
         (csv-data (cdr csv-data))
         (allow-dupes bbdb-allow-duplicates)
         csv-record rd assoc-plus map-bbdb dupes)
    ;; convenient function names
    (fset 'rd 'bbdb-csv-import-rd)
    (fset 'assoc-plus 'bbdb-csv-import-assoc-plus)
    (fset 'map-bbdb (-partial 'bbdb-csv-import-map-bbdb csv-fields))
    ;; we handle duplicates ourselves
    (setq bbdb-allow-duplicates t)
    ;; loop over the csv records
    (while (setq csv-record (map 'list 'cons csv-fields (pop csv-data)))
      (cl-flet*
          ((ca (key list) (cdr (assoc key list))) ;; utility function
           (rd-assoc (root)
                     ;; given ROOT, return a list of data, ignoring empty fields
                     (rd (lambda (elem) (assoc-plus elem csv-record)) (map-bbdb root)))
           (assoc-expand (e)
                         ;; E = data-field-name | (field-name-field data-field)
                         ;; get data from the csv-record and return (field-name data) or nil.
                         (let ((data-name (if (consp e) (ca (car e) csv-record) e))
                               (data (assoc-plus (if (consp e) (cadr e) e) csv-record)))
                           (if data (list data-name data)))))
        ;; set the arguments to bbdb-create-internal, then call it, the end.
        (let ((name (let ((name (rd-assoc :namelist)))
                      ;; prioritize any combination of first middle last over :name
                      (if (>= (length name) 2)
                          (mapconcat 'identity name " ")
                        (car (rd-assoc :name)))))
              (affix (rd-assoc :affix))
              (aka (rd-assoc :aka))
              (organization (rd-assoc :organization))
              (mail (rd-assoc :mail))
              (phone (rd 'vconcat (rd #'assoc-expand (map-bbdb :phone))))
              (address (rd (lambda (e)
                             
                             (let ((al (rd (lambda (elem) ;; al = address lines
                                             (assoc-plus elem csv-record))
                                           (caadr e)))
                                   ;; to use bbdb-csv-import-combined, we can't mapcar
                                   (address-data (--reduce-from (if (member it csv-fields)
                                                                    (cons (ca it csv-record) acc)
                                                                  acc)
                                                                nil (cdadr e)))
                                   (elem-name (car e)))
                               (setq al (nreverse al))
                               (setq address-data (nreverse address-data))
                               ;; make it a list of at least 2 elements
                               (setq al (append al
                                                (-repeat (- 2 (length al)) "")))
                               (when (consp elem-name)
                                 (setq elem-name (ca (car elem-name) csv-record)))
                               
                               ;; determine if non-nil and put together the  minimum set
                               (when (or (not (--all? (zerop (length it)) address-data))
                                         (not (--all? (zerop (length it)) al)))
                                 (when (> 2 (length al))
                                   (setcdr (max 2 (nthcdr (--find-last-index (not (null it))
                                                                             al)
                                                          al)) nil))
                                 (vconcat (list elem-name) (list al) address-data))))
                           (map-bbdb :address)))
              (xfields (rd (lambda (list)
                             (let ((e (car list)))
                               (while (string-match "-" e)
                                 (setq e (replace-match "" nil nil e)))
                               (while (string-match " +" e)
                                 (setq e (replace-match "-" nil nil e)))
                               (setq e (make-symbol (downcase e)))
                               (cons e (cadr list)))) ;; change from (a b) to (a . b)
                           (rd #'assoc-expand (map-bbdb :xfields)))))
          ;; we copy and subvert bbdb's duplicate detection instead of catching
          ;; errors so that we don't interfere with other errors, and can print
          ;; them nicely at the end.
          (let (found-dupe)
            (dolist (elt mail)
              (when (bbdb-gethash elt '(mail))
                (push elt dupes)
                (setq found-dupe t)))
            (when (or allow-dupes (not found-dupe))
              (bbdb-create-internal name affix aka organization mail phone address xfields t))))))
    (when dupes (if allow-dupes
                    (message "Warning, contacts with duplicate email addresses were imported:\n%s" dupes)
                  (message "Skipped contacts with duplicate email addresses:\n%s" dupes)))
    (setq bbdb-allow-duplicates allow-dupes)))

(defun bbdb-csv-import-rd (func list)
  "like mapcar but don't build nil results into the resulting list"
  (--reduce-from (let ((funcreturn (funcall func it)))
                   (if funcreturn
                       (cons funcreturn acc)
                     acc))
                 nil list))

(defun bbdb-csv-import-assoc-plus (key list)
  "Like (cdr assoc ...) but turn an empty string result to nil."
  (let ((result (cdr (assoc key list))))
    (when (not (string= "" result))
      result)))

(provide 'bbdb-csv-import)

;;; bbdb-csv-import.el ends here
