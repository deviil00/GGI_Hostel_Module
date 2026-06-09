from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views
from django.conf import settings
from django.conf.urls.static import static
from django.views.generic import RedirectView
from apps.hostel import views

urlpatterns = [
    path('admin/', admin.site.urls),

    # auth
    path('login/',    auth_views.LoginView.as_view(template_name='base/login.html'), name='login'),
    path('logout/',   auth_views.LogoutView.as_view(), name='logout'),
    path('register/', views.student_register, name='student_register'),

    # dashboard (redirects by role)
    path('', views.dashboard, name='dashboard'),
    path('dashboard/',         views.admin_dashboard_v2, name='admin_dashboard'),
    path('student/dashboard/', views.student_dashboard,  name='student_dashboard'),

    # students
    path('students/',                  views.student_list,   name='student_list'),
    path('students/import/',           views.import_students,       name='import_students'),
    path('students/import/sample/',    views.student_import_sample, name='student_import_sample'),
    path('students/<uuid:pk>/',        views.student_detail, name='student_detail'),
    path('students/<uuid:pk>/set-flag/',            views.student_set_flag,              name='student_set_flag'),
    path('students/<uuid:pk>/toggle-edit-perm/',    views.student_toggle_edit_permission, name='student_toggle_edit_permission'),
    path('students/<uuid:student_pk>/allocate/', views.allocate_room, name='allocate_room'),

    # hostels (admin only)
    path('hostels/',                                          views.hostel_list,        name='hostel_list'),
    path('hostels/bulk-upload/',                              views.bulk_upload_rooms,  name='bulk_upload_rooms'),
    path('hostels/bulk-upload/sample/',                       views.bulk_upload_sample, name='bulk_upload_sample'),
    path('hostels/create/',                                   views.create_hostel,      name='create_hostel'),
    path('hostels/<uuid:pk>/',                                views.hostel_detail,  name='hostel_detail'),
    path('hostels/<uuid:pk>/edit/',                           views.edit_hostel,    name='edit_hostel'),
    path('hostels/<uuid:pk>/delete/',                         views.delete_hostel,  name='delete_hostel'),
    path('hostels/warden/create/',                            views.create_warden,  name='create_warden'),
    path('hostels/<uuid:hostel_pk>/floor/<int:floor_num>/',   views.floor_detail,   name='floor_detail'),
    path('rooms/<uuid:pk>/detail/',                           views.room_detail,    name='room_detail'),

    # rooms
    path('rooms/',                            views.room_list,       name='room_list'),
    path('allocate/',                         views.allocate_room,   name='allocate_room_blank'),
    path('checkout/<uuid:allocation_pk>/',    views.checkout_student, name='checkout_student'),

    # complaints
    path('complaints/',                        views.complaint_list,    name='complaint_list'),
    path('complaints/<uuid:pk>/',              views.complaint_detail,  name='complaint_detail'),
    path('complaints/<uuid:pk>/action/',       views.complaint_action,  name='complaint_action'),
    path('complaints/raise/',                  views.raise_complaint,   name='raise_complaint'),
    path('my/complaints/',                     views.my_complaints,     name='my_complaints'),

    # maintenance incharge
    path('maintenance/dashboard/', views.maintenance_incharge_dashboard, name='maintenance_incharge_dashboard'),

    # fees
    path('fees/',                      views.fee_list,     name='fee_list'),
    path('fees/<uuid:pk>/update/',     views.update_fee,   name='update_fee'),
    path('fees/generate/',             views.generate_fees, name='generate_fees'),

    # attendance
    path('attendance/',                              views.attendance_list,           name='attendance_list'),
    path('attendance/mark/',                         views.mark_attendance,           name='mark_attendance'),
    path('attendance/report/',                       views.attendance_report,         name='attendance_report'),
    path('attendance/toggle/',                       views.toggle_attendance,         name='toggle_attendance'),
    path('attendance/hostel/<uuid:hostel_pk>/',      views.hostel_attendance_detail,  name='hostel_attendance_detail'),

    # forms (registration) — admin
    path('forms/',                                    views.registration_form_list,   name='registration_form_list'),
    path('forms/create/',                             views.create_registration_form, name='create_registration_form'),
    path('forms/<uuid:pk>/edit/',                     views.edit_registration_form,   name='edit_registration_form'),
    path('forms/<uuid:pk>/submissions/',              views.form_submissions,         name='form_submissions'),
    path('forms/<uuid:pk>/analytics/',                views.form_analytics,           name='form_analytics'),
    path('forms/<uuid:pk>/export/',                   views.export_form_excel,        name='export_form_excel'),
    path('forms/submissions/<uuid:pk>/',              views.submission_detail,        name='submission_detail'),
    # backward-compat — old /registration/* paths redirect to /forms/*
    path('registration/',                views.registration_form_list,   name='registration_form_list_compat'),
    path('registration/create/',         views.create_registration_form),
    path('registration/<uuid:pk>/edit/', views.edit_registration_form),
    path('registration/<uuid:pk>/submissions/', views.form_submissions),
    path('registration/submissions/<uuid:pk>/', views.submission_detail),

    # student profile edit + room transfer
    path('students/<uuid:pk>/edit/',     views.edit_student,  name='edit_student'),
    path('students/<uuid:pk>/transfer/', views.transfer_room, name='transfer_room'),

    # reports
    path('fees/defaulters/',                             views.fee_defaulters,          name='fee_defaulters'),
    path('reports/occupancy/',                           views.occupancy_report,        name='occupancy_report'),
    path('reports/occupancy/hostel/<uuid:hostel_pk>/',   views.hostel_occupancy_detail, name='hostel_occupancy_detail'),
    path('reports/gatepasses/', views.gate_pass_report, name='gate_pass_report'),
    path('reports/leaves/',     views.leave_report,     name='leave_report'),
    path('reports/daily/',      RedirectView.as_view(url='/reports/gatepasses/', query_string=False), name='daily_report'),
    path('outsiders/',         views.outsiders,         name='outsiders'),

    # room amenities
    path('amenities/',                          views.amenity_list,    name='amenity_list'),
    path('amenities/<uuid:pk>/delete/',         views.delete_amenity,  name='delete_amenity'),
    path('rooms/<uuid:pk>/edit/',               views.edit_room,       name='edit_room'),

    # announcements
    path('announcements/',                      views.announcement_list,     name='announcement_list'),
    path('announcements/<uuid:pk>/toggle/',     views.toggle_announcement,   name='toggle_announcement'),
    path('announcements/<uuid:pk>/delete/',     views.delete_announcement,   name='delete_announcement'),

    path('my/announcements/', views.student_announcements, name='student_announcements'),

    # registration forms — student
    path('my/forms/',                                      views.available_forms,  name='available_forms'),
    path('my/forms/<uuid:form_pk>/apply/',                 views.apply_form,       name='apply_form'),
    path('my/applications/',                               views.my_applications,  name='my_applications'),

    # student profile
    path('my/profile/',       views.my_profile,            name='my_profile'),
    path('my/profile/edit/',  views.student_edit_profile,  name='student_edit_profile'),

    # gate passes — student
    path('my/gatepasses/',                              views.my_gate_passes,    name='my_gate_passes'),
    path('my/gatepasses/apply/',                        views.apply_gate_pass,   name='apply_gate_pass'),
    path('my/gatepasses/<uuid:pk>/exit-qr/',            views.generate_exit_qr,  name='generate_exit_qr'),
    path('my/gatepasses/<uuid:pk>/entry-qr/',           views.generate_entry_qr, name='generate_entry_qr'),

    # gate passes — admin
    path('gatepasses/',             views.gate_pass_list,   name='gate_pass_list'),
    path('gatepasses/<uuid:pk>/',   views.gate_pass_detail, name='gate_pass_detail'),

    # security guard module
    path('security/',               views.security_dashboard,     name='security_dashboard'),
    path('security/scan/',          views.scan_qr,                name='scan_qr'),
    path('security/verify/',        views.verify_qr,              name='verify_qr'),
    path('security/permit-exit/',   views.guard_permit_exit,      name='guard_permit_exit'),
    path('security/permit-entry/',  views.guard_permit_entry,     name='guard_permit_entry'),
    path('security/allow-exit/',    views.guard_allow_exit,       name='guard_allow_exit'),
    path('security/mark-entry/',    views.guard_mark_entry,       name='guard_mark_entry'),
    path('security/guard/create/',  views.create_security_guard,  name='create_security_guard'),

    # notifications
    path('notifications/',                          views.notification_list,          name='notification_list'),
    path('notifications/clear/',                    views.clear_all_notifications,    name='clear_all_notifications'),
    path('notifications/poll/',                     views.poll_notifications,         name='poll_notifications'),

    # visitors
    path('visitors/',                               views.visitor_list,          name='visitor_list'),
    path('visitors/<uuid:pk>/action/',              views.visitor_action,        name='visitor_action'),
    path('visitors/<uuid:pk>/qr/',                  views.visitor_hostel_qr,     name='visitor_hostel_qr'),
    path('visitors/qr/',                            views.visitor_general_qr,    name='visitor_general_qr'),
    path('visitors/register/',                      views.register_visitor,      name='register_visitor'),
    path('my/visitors/',                            views.my_visitors,           name='my_visitors'),
    path('my/visitors/<uuid:pk>/acknowledge/',      views.visitor_acknowledge,   name='visitor_acknowledge'),
    # Public QR walk-in form (no login)
    path('visitor/walk-in/',                        views.visitor_walkin,        name='visitor_walkin'),
    path('visitor/walk-in/<uuid:hostel_pk>/',       views.visitor_walkin,        name='visitor_walkin_hostel'),

    # mess
    path('mess/',                                   views.mess_management,         name='mess_management'),
    path('mess/menu/',                              views.mess_menu_admin,         name='mess_menu_admin'),
    path('mess/bulk-upload/',                       views.mess_bulk_upload,        name='mess_bulk_upload'),
    path('mess/bulk-upload/sample/',                views.mess_bulk_sample,        name='mess_bulk_sample'),
    path('mess/incharge/',                          views.mess_incharge_dashboard, name='mess_incharge_dashboard'),
    path('mess/feedback/',                          views.mess_feedback_list,      name='mess_feedback_list'),
    path('my/mess/',                                views.mess_menu_student,       name='mess_menu_student'),
    path('superadmin/mess/',                        views.superadmin_mess,         name='superadmin_mess'),

    # discipline
    path('discipline/',                             views.discipline_list,          name='discipline_list'),
    path('discipline/add/',                         views.discipline_add,           name='discipline_add'),
    path('discipline/<uuid:pk>/',                   views.discipline_detail,        name='discipline_detail'),
    path('discipline/<uuid:pk>/action/',            views.discipline_action_view,   name='discipline_action_view'),
    path('discipline/report/',                      views.discipline_report,        name='discipline_report'),
    path('discipline/categories/',                  views.discipline_category_list, name='discipline_category_list'),

    # maintenance
    path('maintenance/',                            views.maintenance_list,   name='maintenance_list'),

    # leave applications
    path('my/leave/',                              views.my_leaves,      name='my_leaves'),
    path('my/leave/apply/',                        views.apply_leave,    name='apply_leave'),
    path('leave/',                                 views.leave_list,     name='leave_list'),
    path('leave/<uuid:pk>/action/',                views.leave_action,   name='leave_action'),

    # security announcements
    path('security/announcements/',                views.security_announcements, name='security_announcements'),

    # bed management
    path('rooms/<uuid:room_pk>/beds/',              views.bed_manage,         name='bed_manage'),

    # semester & reallotment
    path('semesters/',                              views.semester_list,   name='semester_list'),
    path('semesters/close/',                        views.close_semester,  name='close_semester'),
    path('reallotment/',                            views.reallotment,     name='reallotment'),
    path('reallotment/auto/',                       views.auto_reallot,    name='auto_reallot'),
    path('reallotment/single/',                     views.single_reallot,  name='single_reallot'),

    # superadmin
    path('superadmin/',                               views.superadmin_dashboard,    name='superadmin_dashboard'),
    path('superadmin/permissions/',                   views.superadmin_permissions,  name='superadmin_permissions'),
    path('superadmin/staff/',                         views.staff_accounts,          name='staff_accounts'),
    path('superadmin/staff-directory/',               views.staff_directory,         name='staff_directory'),
    path('superadmin/staff-directory/<uuid:uid>/',    views.staff_profile_detail,    name='staff_profile_detail'),
    path('staff/dashboard/',                          views.staff_dashboard,         name='staff_dashboard'),
    path('staff/profile/',                            views.my_staff_profile,        name='my_staff_profile'),

    # ── Assets Management ────────────────────────────────────────────────────
    path('assets/',                                    views.asset_dashboard,        name='asset_dashboard'),
    path('assets/create/',                             views.asset_create,           name='asset_create'),
    path('assets/bulk-upload/',                        views.asset_bulk_upload,      name='asset_bulk_upload'),
    path('assets/bulk-upload/sample/',                 views.asset_bulk_sample,      name='asset_bulk_sample'),
    path('assets/types/',                              views.asset_type_list,        name='asset_type_list'),
    path('assets/<uuid:pk>/',                          views.asset_detail,           name='asset_detail'),
    path('assets/<uuid:pk>/edit/',                     views.asset_edit,             name='asset_edit'),
    path('assets/<uuid:pk>/discard/',                  views.asset_discard,          name='asset_discard'),
    path('assets/<uuid:pk>/transfer/',                 views.asset_transfer_request, name='asset_transfer_request'),
    path('assets/transfers/<uuid:transfer_pk>/action/', views.asset_transfer_action, name='asset_transfer_action'),

    # ── Gate Pass Categories & Academic Calendar ──────────────────────────────
    path('superadmin/gp-categories/',                  views.gp_category_list,       name='gp_category_list'),
    path('superadmin/gp-categories/<uuid:pk>/toggle/', views.gp_category_toggle,     name='gp_category_toggle'),
    path('superadmin/calendar/',                       views.academic_calendar,      name='academic_calendar'),

    # ── Leave Categories & Extension ──────────────────────────────────────────
    path('superadmin/academic-calendar/sample/',       views.academic_calendar_sample, name='academic_calendar_sample'),
    path('superadmin/leave-categories/',               views.leave_category_list,    name='leave_category_list'),
    path('superadmin/leave-categories/<uuid:pk>/toggle/', views.leave_category_toggle, name='leave_category_toggle'),
    path('my/leave/<uuid:leave_pk>/extend/',           views.leave_extension_request, name='leave_extension_request'),
    path('leave/extension/<uuid:ext_pk>/action/',      views.leave_extension_action, name='leave_extension_action'),
    path('leave/<uuid:pk>/warden-action/',             views.leave_warden_action,    name='leave_warden_action'),
    path('leave/hod/',                                 views.hod_leave_list,         name='hod_leave_list'),
    path('leave/<uuid:pk>/hod-action/',                views.leave_hod_action,       name='leave_hod_action'),
    path('superadmin/restrictions/',                   views.restriction_list,        name='restriction_list'),
    path('superadmin/restrictions/add/',               views.restriction_add,         name='restriction_add'),
    path('superadmin/restrictions/<uuid:pk>/lift/',    views.restriction_lift,        name='restriction_lift'),
    path('superadmin/quota-policies/',                 views.quota_policy_list,       name='quota_policy_list'),

    # ── Download Reports ─────────────────────────────────────────────────────
    path('reports/download/',                          views.reports_download_center,  name='reports_download_center'),
    path('reports/room-transfer/',                     views.report_room_transfer,     name='report_room_transfer'),
    path('reports/attendance/',                        views.report_attendance,        name='report_attendance'),
    path('reports/fees/',                              views.report_fees,              name='report_fees'),
    path('reports/gatepass/',                          views.report_gatepass,          name='report_gatepass'),
    path('reports/mess/',                              views.report_mess,              name='report_mess'),
    path('reports/hostel/<uuid:hostel_pk>/daily/',     views.report_hostel_daily,      name='report_hostel_daily'),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    urlpatterns += static(settings.STATIC_URL, document_root=settings.STATIC_ROOT)
