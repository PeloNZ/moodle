<?php

// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

require_once($CFG->dirroot.'/grade/export/lib.php');

class grade_export_xls extends grade_export {

    public $plugin = 'xls';
    protected $showgroups = false;
    protected $showcohorts = false;

    /**
     * To be implemented by child classes
     */
    public function print_grades() {
        global $CFG;
        require_once($CFG->dirroot.'/lib/excellib.class.php');

        $export_tracking = $this->track_exports();

        $strgrades = get_string('grades');

        $gui = new graded_users_iterator($this->course, $this->columns, $this->groupid, 'lastname', 'ASC','firstname', 'ASC', $this->showcohorts, $this->showgroups);
        $gui->require_active_enrolment($this->onlyactive);
        $gui->allow_user_custom_fields($this->usercustomfields);
        $gui->init();
        $groups = array();
        if (!empty($gui->group_rs)) { //if showgroups is set, iterate over all group info to make it useful
            foreach ($gui->group_rs as $gr) {
                if (empty($groups[$gr->id])) {
                    $groups[$gr->id] = $gr->name;
                } else {
                    $groups[$gr->id] .= ', '.$gr->name;
                }
            }
        }
        $cohorts = array();
        if (!empty($gui->cohorts_rs)) { //if showcohorts is set, iterate over all cohort info to make it useful
            foreach ($gui->cohorts_rs as $cr) {
                if (empty($cohorts[$cr->id])) {
                    $cohorts[$cr->id] = $cr->name;
                } else {
                    $cohorts[$cr->id] .= ', '.$cr->name;
                }
            }
        }
        // Calculate file name
        $shortname = format_string($this->course->shortname, true, array('context' => context_course::instance($this->course->id)));
        $downloadfilename = clean_filename("$shortname $strgrades.xls");
        // Creating a workbook
        $workbook = new MoodleExcelWorkbook("-");
        // Sending HTTP headers
        $workbook->send($downloadfilename);
        // Adding the worksheet
        $myxls = $workbook->add_worksheet($strgrades);

        // Print names of all the fields
        $profilefields = grade_helper::get_user_profile_fields($this->course->id, $this->usercustomfields);
        foreach ($profilefields as $id => $field) {
            $myxls->write_string(0, $id, $field->fullname);
        }
        $pos = count($profilefields);

        if (!empty($gui->group_rs)) {
            $myxls->write_string(0, $pos++,get_string("group"));
        }
        if (!empty($gui->cohorts_rs)) {
            $myxls->write_string(0, $pos++,get_string('cohorts', 'cohort'));
        }

        foreach ($this->columns as $grade_item) {
            $myxls->write_string(0, $pos++, $this->format_column_name($grade_item));
            $myxls->write_string(0, $pos++, get_string('time'));

            // Add a column_feedback column
            if ($this->export_feedback) {
                $myxls->write_string(0, $pos++, $this->format_column_name($grade_item, true));
            }
        }

        // Print all the lines of data.
        $i = 0;
        $geub = new grade_export_update_buffer();
        while ($userdata = $gui->next_user()) {
            $i++;
            $user = $userdata->user;

            foreach ($profilefields as $id => $field) {
                $fieldvalue = grade_helper::get_user_field_value($user, $field);
                $myxls->write_string($i, $id, $fieldvalue);
            }
            $j = count($profilefields);

            if (!empty($gui->group_rs)) {
                if (!empty($groups[$user->id])) {
                    $myxls->write_string($i,$j,$groups[$user->id]);
                }
                $j++;
            }
            if (!empty($gui->cohorts_rs)) {
                if (!empty($cohorts[$user->id])) {
                    $myxls->write_string($i,$j,$cohorts[$user->id]);
                }
                $j++;
            }
            foreach ($userdata->grades as $itemid => $grade) {
                if ($export_tracking) {
                    $status = $geub->track($grade);
                }

                $gradestr = $this->format_grade($grade);
                if (is_numeric($gradestr)) {
                    $myxls->write_number($i,$j++,$gradestr);
                }
                else {
                    $myxls->write_string($i,$j++,$gradestr);
                }
                //write timestamp for grade
                $myxls->write_string($i,$j++,$userdata->gradetimes[$itemid]);

                // writing feedback if requested
                if ($this->export_feedback) {
                    $myxls->write_string($i, $j++, $this->format_feedback($userdata->feedbacks[$itemid]));
                }
            }
        }
        $gui->close();
        $geub->close();

    /// Close the workbook
        $workbook->close();

        exit;
    }
    public function grade_export_xls($course, $groupid=0, $itemlist='', $export_feedback=false, $updatedgradesonly = false, $displaytype = GRADE_DISPLAY_TYPE_REAL, $decimalpoints = 2, $showgroups = false, $showcohorts = false) {
        $this->showgroups = $showgroups;
        $this->showcohorts = $showcohorts;
        parent::grade_export($course, $groupid, $itemlist, $export_feedback, $updatedgradesonly, $displaytype, $decimalpoints);
    }

    /**
     * Returns array of parameters used by dump.php and export.php.
     * @return array
     */
    public function get_export_params() {
            $itemids = array_keys($this->columns);
            $itemidsparam = implode(',', $itemids);
        if (empty($itemidsparam)) {
                    $itemidsparam = '-1';
        }

        $params = array('id'                =>$this->course->id,
                'groupid'           =>$this->groupid,
                'itemids'           =>$itemidsparam,
                'export_letters'    =>$this->export_letters,
                'export_feedback'   =>$this->export_feedback,
                'updatedgradesonly' =>$this->updatedgradesonly,
                'displaytype'       =>$this->displaytype,
                'decimalpoints'     =>$this->decimalpoints,
                'showgroups'        =>$this->showgroups,
                'showcohorts'       =>$this->showcohorts);

        return $params;
    }

}


