;;; bbdb3-csv-import.el  --- import csv to bbdb version 3+ -*- lexical-binding: t -*-

;; Copyright (C) 2014 by Ian Kelling

;; Author: Ian Kelling <ian@iankelling.org>
;; Created: 1 Apr 2014
;; Version: 1.0
;; Keywords: csv, util, bbdb

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
;; version 3+. Programs such as Thunderbird, Gmail, Linkedin, and Outlook allow
;; for exporting contact data as csv files. See ASynK for syncing bbdb/google/outlook.

;;; Installation:
;;
;; dependencies: pcsv.el, dash.el, bbdb
;; These are available via marmalade/melpa or the internet
;;
;; Add to init file or execute manually as this may be a one time usage:
;; (load-file FILENAME-OF-THIS-FILE)
;; or
;; (add-to-list 'load-path DIRECTORY-CONTAINING-THIS-FILE)
;; (require 'bbdb3-csv-import)

;;; Usage:
;;
;; Backup or rename any existing ~/.bbdb and ~/.emacs.d/bbdb while testing that
;; the import works correctly.
;;
;; Assign a mapping table. Predefined ones listed here:
;; (setq bbdb3-csv-import-mapping-table bbdb3-csv-import-thunderbird)
;; (setq bbdb3-csv-import-mapping-table bbdb3-csv-import-gmail)
;; (setq bbdb3-csv-import-mapping-table bbdb3-csv-import-linkedin)
;; (setq bbdb3-csv-import-mapping-table bbdb3-csv-import-outlook-web)
;; 
;;
;; Simply call `bbdb3-csv-import-buffer' or
;; `bbdb3-csv-import-file'. Interactively they prompt for file/buffer. Use
;; non-interactively for no prompts.
;;
;; If you need to define your own mapping table, it should not be too hard.  Use
;; the existing tables as an example, and perhaps the test data within this
;; project. Please send any new mapping tables to the maintainer listed in this
;; file. The maintainer should be able to help with any issues and may create a
;; new mapping table given sample data.
;;
;; Tips for testing:
;; - bbdb doesn't work if you delete the bbdb database file in
;;   the middle of an emacs session. If you want to empty the current bbdb database,
;;   do M-x bbdb then .* then C-u * d on the beginning of a record.
;; - After changing a mapping table, don't forget to re-execute
;;   (setq bbdb3-csv-import-mapping-table ...) so that it propagates.
;;
;; Todo: It would be nice if we would programatically or manually merge all the
;; mapping tables, then we would not have to set one.

(require 'pcsv)
(require 'dash)
(require 'bbdb-com)
(eval-when-compile (require 'cl))

(defconst bbdb3-csv-import-thunderbird
  '(("firstname" "First Name")
    ("lastname" "Last Name")
    ("name" "Display Name")
    ("aka" "Nickname")
    ("mail" "Primary Email" "Secondary Email")
    ("phone" "Work Phone" "Home Phone" "Fax Number" "Pager Number" "Mobile Number")
    ("address"
     (("home address"
       (("Home Address" "Home Address 2")
        "Home City" "Home State"
        "Home ZipCode" "Home Country"))
      ("work address"
       (("Work Address" "Work Address 2")
        "Work City" "Work State"
        "Work ZipCode" "Work Country"))))
    ("organization" "Organization")
    ("xfields" "Web Page 1" "Web Page 2" "Birth Year" "Birth Month"
     "Birth Day" "Department" "Custom 1" "Custom 2" "Custom 3"
     "Custom 4" "Notes" "Job Title"))
  "Thunderbird csv format")

(defconst bbdb3-csv-import-linkedin
  '(("firstname" "First Name")
    ("lastname" "Last Name")
    ("middlename" "Middle Name")
    ("mail" "E-mail Address" "E-mail 2 Address" "E-mail 3 Address")
    ("phone"
     "Assistant's Phone" "Business Fax" "Business Phone"
     "Business Phone 2" "Callback" "Car Phone"
     "Company Main Phone" "Home Fax" "Home Phone"
     "Home Phone 2" "ISDN" "Mobile Phone"
     "Other Fax" "Other Phone" "Pager"
     "Primary Phone" "Radio Phone" "TTY/TDD Phone" "Telex")
    ("address"
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
    ("organization" "Company")
    ("xfields"
     "Suffix" "Department" "Job Title" "Assistant's Name"
     "Birthday" "Manager's Name" "Notes" "Other Address PO Box"
     "Spouse" "Web Page" "Personal Web Page"))
  "Linkedin export in the Outlook csv format.")


;; note. PO Box and Extended Address are added as additional address street lines if they exist.
;; If you don't like this, just delete them from this fiel.
;; If you want some other special handling, it will need to be coded.
(defconst bbdb3-csv-import-gmail
  '(("firstname" "Given Name")
    ("lastname" "Family Name")
    ("name" "Name")
    ("mail" (repeat "E-mail 1 - Value"))
    ("phone" (repeat ("Phone 1 - Type" "Phone 1 - Value")))
    ("address"
     (repeat (("Address 1 - Type")
              (("Address 1 - Street" "Address 1 - PO Box" "Address 1 - Extended Address")
               "Address 1 - City" "Address 1 - Region"
               "Address 1 - Postal Code" "Address 1 - Country"))))
    ("organization" (repeat "Organization 1 - Name"))
    ("xfields"
     "Additional Name" "Yomi Name" "Given Name Yomi"
     "Additional Name Yomi" "Family Name Yomi" "Name Prefix"
     "Name Suffix" "Initials" "Nickname"
     "Short Name" "Maiden Name" "Birthday"
     "Gender" "Location" "Billing Information"
     "Directory Server" "Mileage" "Occupation"
     "Hobby" "Sensitivity" "Priority"
     "Subject" "Notes" "Group Membership"
     ;; Gmail wouldn't let me add more than 1 organization, but no harm in
     ;; looking for multiple since the field name implies the possibility.
     (repeat  
      "Organization 1 - Type" "Organization 1 - Yomi Name"
      "Organization 1 - Title" "Organization 1 - Department"
      "Organization 1 - Symbol" "Organization 1 - Location"
      "Organization 1 - Job Description")
     (repeat ("Relation 1 - Type" "Relation 1 - Value"))
     (repeat ("Website 1 - Type" "Website 1 - Value"))
     (repeat ("Event 1 - Type" "Event 1 - Value"))
     (repeat ("Custom Field 1 - Type" "Custom Field 1 - Value"))))
  "Gmail csv export format")


(defconst bbdb3-csv-import-gmail-typed-email
  (append  (car (last bbdb3-csv-import-gmail)) '((repeat "E-mail 1 - Type")))
  "Like the first Gmail mapping, but use custom fields to store
   Gmail's email labels. This is separate because I assume most
   people don't use those labels and using the default labels
   would create useless custom fields.")

(defconst bbdb3-csv-import-outlook-typed-email
  (append  (car (last bbdb3-csv-import-outlook-web)) '((repeat "E-mail 1 - Type")))
  "Like the previous var, but for outlook-web.
Adds email labels as custom fields.")


(defconst bbdb3-csv-import-outlook-web
  '(("firstname" "Display Name" "First Name")
    ("lastname" "Last Name")
    ("middlename" "Middle Name")
    ("mail" "E-mail Address" "E-mail 2 Address" "E-mail 3 Address")
    ("phone"
     "Assistant's Phone" "Business Fax" "Business Phone"
     "Business Phone 2" "Callback" "Car Phone"
     "Company Main Phone" "Home Fax" "Home Phone"
     "Home Phone 2" "ISDN" "Mobile Phone"
     "Other Fax" "Other Phone" "Pager"
     "Primary Phone" "Radio Phone" "TTY/TDD Phone" "Telex")
    ("address"
     (("business address"
       (("Business Street")
        "Business City" "Business State"
        "Business Postal Code" "Business Country"))
      ("home address"
       (("Home Street")
        "Home City" "Home State"
        "Home Postal Code" "Home Country"))
      ("other address"
       (("Other Street" "")
        "Other City" "Other State"
        "Other Postal Code" "Other Country"))))
    ("organization" "Company")
    ("xfields"
     "Anniversary" "Family Name Yomi" "Given Name Yomi"
     "Suffix" "Department" "Job Title" "Birthday" "Manager's Name" "Notes"
     "Spouse" "Web Page"))
  "Hotmail.com, outlook.com, live.com, etc.
Based on 'Export for outlook.com and other services',
not the export for Outlook 2010 and 2013.")

;(defconst bbdb3-csv-import-combined)


(defvar bbdb3-csv-import-mapping-table nil
  "The table which maps bbdb3 fields to csv fields.
Use the default as an example to map non-thunderbird data.
Name used is firstname + lastname or name.
After the car, all names should map to whatever csv
field names are used in the first row of csv data.
Many fields are optional. If you aren't sure if one is,
best to just try it. The doc string for `bbdb-create-internal'
may be useful for determining which fields are required.")

;;;###autoload
(defun bbdb3-csv-import-file (filename)
  "Parse and import csv file FILENAME to bbdb3."
  (interactive "fCSV file containg contact data: ")
  (bbdb3-csv-import-buffer (find-file-noselect filename)))


;;;###autoload
(defun bbdb3-csv-import-buffer (&optional buffer-or-name) 
  "Parse and import csv BUFFER-OR-NAME to bbdb3.
Argument is a buffer or name of a buffer.
Defaults to current buffer."
  (interactive "bBuffer containing CSV contact data: ")
  (when (null bbdb3-csv-import-mapping-table)
    (error "error: `bbdb3-csv-import-mapping-table' is nil. Please set it and rerun."))
  (let* ((csv-fields (pcsv-parse-buffer (get-buffer (or buffer-or-name (current-buffer)))))
         (csv-contents (cdr csv-fields))
         (csv-fields (car csv-fields))
         (initial-duplicate-value bbdb-allow-duplicates)
         csv-record rd assoc-plus flatten1)
    ;; convenient function names
    (fset 'rd 'bbdb3-csv-import-rd)
    (fset 'assoc-plus 'bbdb3-csv-import-assoc-plus)
    (fset 'flatten1 'bbdb3-csv-import-flatten1)
    ;; Easier to allow duplicates and handle them post import vs failing as
    ;; soon as we find one.
    (setq bbdb-allow-duplicates t)
    ;; loop over the csv records
    (while (setq csv-record (map 'list 'cons csv-fields (pop csv-contents)))
      (cl-flet*
          ((expand-repeats (list)
                           ;; return new list where elements from LIST in form
                           ;; (repeat elem1 ...) become ((elem1 ...) [(elem2 ...)] ...)
                           ;; For as many repeating numbered fields exist in the csv fields.
                           ;; elem can be a string or a tree (a list with possibly lists inside it)
                           (--reduce-from (if (not (and (consp it) (eq (car it) 'repeat)))
                                              (cons it acc)
                                            (setq it (cdr it))
                                            (let* ((i 1)
                                                   (first-field (car (flatten it))))
                                              (setq acc (cons it acc))
                                              ;; use first-field to test if there is another repetition.
                                              (while (member (replace-num (setq i (1+ i)) first-field) csv-fields)
                                                (cl-labels ((fun (cell)
                                                                 (if (consp cell)
                                                                     (mapcar #'fun cell)
                                                                   (replace-num i cell))))
                                                  (setq acc (cons (fun it) acc))))
                                              acc))
                                          nil list))
           (map-bbdb3 (root)
                      ;; ROOT = a root element from bbdb3-csv-import-mapping-table.
                      ;; Get the actual csv-fields, including variably repeated ones. flattened
                      ;; by one because repeated fields are put in sub-lists, but
                      ;; after expanding them, that extra depth is no longer
                      ;; useful. Small quirk: address mappings without 'repeat
                      ;; need to be grouped in a list because they contain sublists that we
                      ;; don't want flattened. Better this than more complex code.
                      (flatten1 (expand-repeats (cdr (assoc root bbdb3-csv-import-mapping-table)))))
           (rd-assoc (root)
                     ;; given ROOT, return a list of data, ignoring empty fields
                     (rd (lambda (elem) (assoc-plus elem csv-record)) (map-bbdb3 root)))
           (mapcar-assoc (list)
                         ;; given LIST of fields,return a list of data with nil in place of an empty field
                         (mapcar (lambda (elem) (cdr (assoc elem csv-record))) list))
           (assoc-expand (e)
                         ;; E = data-field-name | (field-name-field data-field)
                         ;; get data from the csv-record and return
                         ;; (field-name data) or nil.
                         (let ((data-name (if (consp e) (cdr (assoc (car e) csv-record)) e))
                               (data (assoc-plus (if (consp e) (cadr e) e) csv-record)))
                           (if data (list data-name data))))
           (replace-num (num string)
                        ;; in STRING, replace all groups of numbers with NUM
                        (replace-regexp-in-string "[0-9]+" (number-to-string num) string))
           (map-assoc (field)
                      ;; For mappings with just 1 simple csv-field, get it's data
                      (car (rd-assoc field))))

        (let ((name (let ((first (map-assoc "firstname"))
                          (middle (map-assoc "middlename"))
                          (last (map-assoc "lastname"))
                          (name (map-assoc "name")))
                      ;; prioritize any combination of first middle last over just "name"
                      (if (or (and first last) (and first middle) (and middle last))
                          ;; purely historical note.
                          ;; using (cons first last) as argument works the same as (concat first " " last)
                          (concat (or first middle) " " (or middle last) (when (and first middle) (concat " " last) ))
                        (or name first middle last ""))))
              (phone (rd 'vconcat (rd #'assoc-expand (map-bbdb3 "phone"))))
              (mail (rd-assoc "mail"))
              (xfields (rd (lambda (list)
                             (let ((e (car list)))
                               (while (string-match "-" e)
                                 (setq e (replace-match "" nil nil e)))
                               (while (string-match " +" e)
                                 (setq e (replace-match "-" nil nil e)))
                               (setq e (make-symbol (downcase e)))
                               (cons e (cadr list)))) ;; change from (a b) to (a . b)
                           (rd #'assoc-expand (map-bbdb3 "xfields"))))
              (address (rd (lambda (mapping-elem)
                             (let ((address-lines (rd (lambda (elem)
                                                        (assoc-plus elem csv-record))
                                                      (caadr mapping-elem)))
                                   (address-data (mapcar-assoc (cdadr mapping-elem)))
                                   (elem-name (car mapping-elem)))
                               (if (= (length address-lines) 1)
                                   (setq address-lines (-snoc address-lines "")))
                               (when (consp elem-name)
                                 (setq elem-name (cdr (assoc (car elem-name) csv-record))))
                               
                               ;; determine if non-nil and put together the  minimum set
                               (when (or (not (--all? (zerop (length it)) address-data))
                                         (not (--all? (zerop (length it)) address-lines)))
                                 (when (> 2 (length address-lines))
                                   (setcdr (max 2 (nthcdr (--find-last-index (not (null it))
                                                                             address-lines)
                                                          address-lines)) nil))
                                 (vconcat (list elem-name) (list address-lines) address-data))))
                           (map-bbdb3 "address")))
              (organization (rd-assoc "organization"))
              (affix (map-assoc "affix"))
              (aka (rd-assoc "aka")))
          (bbdb-create-internal name affix aka organization mail phone address xfields t))))
    (setq bbdb-allow-duplicates initial-duplicate-value)))

(defun bbdb3-csv-import-flatten1 (list)
  "flatten LIST by 1 level."
  (--reduce-from (if (consp it)
                     (-concat acc it)
                   (-snoc acc it))
                 nil list))

;;;###autoload
(defun bbdb3-csv-import-rd (func list)
  "like mapcar but don't build nil results into the resulting list"
  (--reduce-from (let ((funcreturn (funcall func it)))
                   (if funcreturn
                       (cons funcreturn acc)
                     acc))
                 nil list))

;;;###autoload
(defun bbdb3-csv-import-assoc-plus (key list)
  "Like `assoc' but turn an empty string result to nil."
  (let ((result (cdr (assoc key list))))
    (when (not (string= "" result))
      result)))



(provide 'bbdb3-csv-import)

;;; bbdb3-csv-import.el ends here
