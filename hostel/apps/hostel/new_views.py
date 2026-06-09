
# ═══════════════════════════════════════════════════════════════════════════════
# ASSETS MANAGEMENT MODULE
# ═══════════════════════════════════════════════════════════════════════════════

@login_required
def asset_dashboard(request):
    role = request.user.role
    if role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')

    f_hostel    = request.GET.get('hostel', '')
    f_type      = request.GET.get('asset_type', '')
    f_condition = request.GET.get('condition', '')
    f_status    = request.GET.get('status', '')
    f_q         = request.GET.get('q', '').strip()

    qs = Asset.objects.select_related('hostel', 'room', 'asset_type').all()
    if f_hostel:    qs = qs.filter(hostel_id=f_hostel)
    if f_type:      qs = qs.filter(asset_type_id=f_type)
    if f_condition: qs = qs.filter(condition=f_condition)
    if f_status:    qs = qs.filter(status=f_status)
    if f_q:
        qs = qs.filter(Q(name__icontains=f_q) | Q(asset_code__icontains=f_q) |
                       Q(vendor_name__icontains=f_q))

    counts = {
        'total':             Asset.objects.count(),
        'active':            Asset.objects.filter(status='active').count(),
        'damaged':           Asset.objects.filter(condition='damaged').count(),
        'maintenance':       Asset.objects.filter(condition='under_maintenance').count(),
        'lost':              Asset.objects.filter(status='lost').count(),
        'pending_transfers': AssetTransfer.objects.filter(status='pending').count(),
    }

    paginator = Paginator(qs, 25)
    page_obj  = paginator.get_page(request.GET.get('page'))
    return render(request, 'admin/assets.html', {
        'page_obj': page_obj, 'counts': counts,
        'hostels': Hostel.objects.all(),
        'asset_types': AssetType.objects.all(),
        'conditions': Asset.Condition.choices,
        'statuses': Asset.Status.choices,
        'filters': {'hostel': f_hostel, 'asset_type': f_type,
                    'condition': f_condition, 'status': f_status, 'q': f_q},
    })


@login_required
def asset_create(request):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        name       = request.POST.get('name', '').strip()
        asset_code = request.POST.get('asset_code', '').strip()
        if not name or not asset_code:
            messages.error(request, 'Name and asset code are required.')
            return redirect('asset_create')
        if Asset.objects.filter(asset_code=asset_code).exists():
            messages.error(request, f'Asset code "{asset_code}" already exists.')
            return redirect('asset_create')
        floor_val = request.POST.get('floor', '')
        asset = Asset.objects.create(
            name=name, asset_code=asset_code,
            asset_type_id=request.POST.get('asset_type') or None,
            hostel_id=request.POST.get('hostel') or None,
            floor=int(floor_val) if floor_val else None,
            room_id=request.POST.get('room') or None,
            condition=request.POST.get('condition', 'good'),
            vendor_name=request.POST.get('vendor_name', '').strip(),
            vendor_contact=request.POST.get('vendor_contact', '').strip(),
            vendor_invoice=request.POST.get('vendor_invoice', '').strip(),
            purchase_date=request.POST.get('purchase_date') or None,
            purchase_cost=request.POST.get('purchase_cost') or None,
            assigned_date=request.POST.get('assigned_date') or None,
            notes=request.POST.get('notes', '').strip(),
            added_by=request.user,
        )
        AssetLog.objects.create(
            asset=asset, action=AssetLog.Action.CREATED,
            description=f'Asset created at {asset.location_display}.',
            performed_by=request.user,
        )
        messages.success(request, f'Asset "{asset.name}" created.')
        return redirect('asset_detail', pk=asset.pk)
    return render(request, 'admin/asset_form.html', {
        'hostels': Hostel.objects.all(), 'asset_types': AssetType.objects.all(),
        'conditions': Asset.Condition.choices,
        'rooms': Room.objects.select_related('hostel').order_by('hostel__name', 'room_number'),
        'action': 'Create',
    })


@login_required
def asset_edit(request, pk):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    asset = get_object_or_404(Asset, pk=pk)
    if request.method == 'POST':
        old_cond = asset.condition
        asset.name           = request.POST.get('name', asset.name).strip()
        asset.asset_type_id  = request.POST.get('asset_type') or None
        asset.hostel_id      = request.POST.get('hostel') or None
        fv = request.POST.get('floor', '')
        asset.floor          = int(fv) if fv else None
        asset.room_id        = request.POST.get('room') or None
        asset.condition      = request.POST.get('condition', asset.condition)
        asset.vendor_name    = request.POST.get('vendor_name', '').strip()
        asset.vendor_contact = request.POST.get('vendor_contact', '').strip()
        asset.vendor_invoice = request.POST.get('vendor_invoice', '').strip()
        asset.purchase_date  = request.POST.get('purchase_date') or None
        pc = request.POST.get('purchase_cost', '')
        asset.purchase_cost  = pc if pc else None
        asset.assigned_date  = request.POST.get('assigned_date') or None
        asset.notes          = request.POST.get('notes', '').strip()
        asset.save()
        action = AssetLog.Action.CONDITION_CHANGED if old_cond != asset.condition else AssetLog.Action.UPDATED
        desc = (f'Condition changed: {old_cond} → {asset.condition}.'
                if action == AssetLog.Action.CONDITION_CHANGED else 'Details updated.')
        AssetLog.objects.create(asset=asset, action=action, description=desc, performed_by=request.user)
        messages.success(request, 'Asset updated.')
        return redirect('asset_detail', pk=asset.pk)
    return render(request, 'admin/asset_form.html', {
        'asset': asset, 'hostels': Hostel.objects.all(),
        'asset_types': AssetType.objects.all(), 'conditions': Asset.Condition.choices,
        'rooms': Room.objects.select_related('hostel').order_by('hostel__name', 'room_number'),
        'action': 'Edit',
    })


@login_required
def asset_detail(request, pk):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    asset = get_object_or_404(
        Asset.objects.select_related('hostel', 'room', 'asset_type', 'added_by', 'discarded_by'), pk=pk)
    logs      = asset.logs.select_related('performed_by').order_by('-created_at')
    transfers = asset.transfers.select_related(
        'from_hostel', 'from_room', 'to_hostel', 'to_room', 'requested_by', 'approved_by').order_by('-created_at')
    return render(request, 'admin/asset_detail.html', {
        'asset': asset, 'logs': logs, 'transfers': transfers,
        'hostels': Hostel.objects.all(),
        'rooms': Room.objects.select_related('hostel').order_by('hostel__name', 'room_number'),
        'conditions': Asset.Condition.choices,
    })


@login_required
@require_POST
def asset_discard(request, pk):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    asset  = get_object_or_404(Asset, pk=pk)
    reason = request.POST.get('reason', '').strip()
    if not reason:
        messages.error(request, 'Discard reason is required.')
        return redirect('asset_detail', pk=pk)
    asset.status = Asset.Status.DISCARDED
    asset.condition = Asset.Condition.DISCARDED
    asset.discard_reason = reason
    asset.discarded_at   = timezone.now()
    asset.discarded_by   = request.user
    asset.save()
    AssetLog.objects.create(asset=asset, action=AssetLog.Action.DISCARDED,
                            description=f'Discarded. Reason: {reason}', performed_by=request.user)
    messages.success(request, f'Asset "{asset.name}" discarded.')
    return redirect('asset_dashboard')


@login_required
def asset_transfer_request(request, pk):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    asset = get_object_or_404(Asset, pk=pk)
    if request.method == 'POST':
        to_hostel_id  = request.POST.get('to_hostel', '')
        to_floor_val  = request.POST.get('to_floor', '')
        to_room_id    = request.POST.get('to_room', '')
        reason        = request.POST.get('reason', '').strip()
        transfer_date = request.POST.get('transfer_date', str(timezone.now().date()))
        if not reason:
            messages.error(request, 'Transfer reason is required.')
            return redirect('asset_detail', pk=pk)
        is_superadmin = request.user.role == User.Role.SUPER_ADMIN
        t_status = AssetTransfer.Status.COMPLETED if is_superadmin else AssetTransfer.Status.PENDING
        transfer = AssetTransfer.objects.create(
            asset=asset,
            from_hostel=asset.hostel, from_room=asset.room, from_floor=asset.floor,
            to_hostel_id=to_hostel_id or None, to_room_id=to_room_id or None,
            to_floor=int(to_floor_val) if to_floor_val else None,
            reason=reason, transfer_date=transfer_date,
            requested_by=request.user, status=t_status,
        )
        if is_superadmin:
            transfer.approved_by = request.user
            transfer.approved_at = timezone.now()
            transfer.save()
            asset.hostel_id = to_hostel_id or None
            asset.room_id   = to_room_id or None
            asset.floor     = int(to_floor_val) if to_floor_val else None
            asset.assigned_date = transfer_date
            asset.save()
            AssetLog.objects.create(asset=asset, action=AssetLog.Action.TRANSFERRED,
                                    description=f'Transferred to {asset.location_display}.', performed_by=request.user)
            messages.success(request, 'Asset transferred.')
        else:
            AssetLog.objects.create(asset=asset, action=AssetLog.Action.TRANSFERRED,
                                    description='Transfer requested — awaiting approval.', performed_by=request.user)
            messages.success(request, 'Transfer request submitted.')
        return redirect('asset_detail', pk=pk)
    return redirect('asset_detail', pk=pk)


@login_required
@require_POST
def asset_transfer_action(request, transfer_pk):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    transfer = get_object_or_404(AssetTransfer, pk=transfer_pk)
    action   = request.POST.get('action', '')
    remarks  = request.POST.get('remarks', '').strip()
    if action == 'approve' and transfer.status == AssetTransfer.Status.PENDING:
        transfer.status = AssetTransfer.Status.COMPLETED
        transfer.approved_by = request.user
        transfer.approved_at = timezone.now()
        transfer.remarks     = remarks
        transfer.save()
        a = transfer.asset
        a.hostel_id = transfer.to_hostel_id
        a.room_id   = transfer.to_room_id
        a.floor     = transfer.to_floor
        a.assigned_date = str(transfer.transfer_date)
        a.save()
        AssetLog.objects.create(asset=a, action=AssetLog.Action.TRANSFERRED,
                                description=f'Transfer approved. Now at {a.location_display}.', performed_by=request.user)
        messages.success(request, 'Transfer approved.')
    elif action == 'reject' and transfer.status == AssetTransfer.Status.PENDING:
        transfer.status  = AssetTransfer.Status.REJECTED
        transfer.remarks = remarks
        transfer.save()
        messages.info(request, 'Transfer rejected.')
    return redirect('asset_detail', pk=transfer.asset_id)


@login_required
def asset_bulk_upload(request):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        csv_file = request.FILES.get('csv_file')
        if not csv_file:
            messages.error(request, 'Please select a CSV file.')
            return redirect('asset_bulk_upload')
        decoded = csv_file.read().decode('utf-8-sig')
        reader  = csv.DictReader(io.StringIO(decoded))
        created = 0
        errors  = []
        for i, row in enumerate(reader, start=2):
            code = row.get('asset_code', '').strip()
            name = row.get('name', '').strip()
            if not code or not name:
                errors.append(f'Row {i}: asset_code and name required.')
                continue
            if Asset.objects.filter(asset_code=code).exists():
                errors.append(f'Row {i}: "{code}" already exists — skipped.')
                continue
            hostel = None
            hn = row.get('hostel', '').strip()
            if hn:
                hostel = Hostel.objects.filter(name__iexact=hn).first()
            at = None
            tn = row.get('asset_type', '').strip()
            if tn:
                at, _ = AssetType.objects.get_or_create(name=tn)
            fv = row.get('floor', '')
            Asset.objects.create(
                asset_code=code, name=name, asset_type=at, hostel=hostel,
                floor=int(fv) if fv else None,
                condition=row.get('condition', 'good'),
                vendor_name=row.get('vendor_name', ''),
                purchase_date=row.get('purchase_date') or None,
                purchase_cost=row.get('purchase_cost') or None,
                notes=row.get('notes', ''),
                added_by=request.user,
            )
            created += 1
        if created:
            messages.success(request, f'{created} asset(s) imported.')
        for err in errors[:5]:
            messages.warning(request, err)
        return redirect('asset_dashboard')
    return render(request, 'admin/asset_bulk_upload.html')


def asset_bulk_sample(request):
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="asset_upload_sample.csv"'
    w = csv.writer(response)
    w.writerow(['asset_code', 'name', 'asset_type', 'hostel', 'floor', 'condition',
                'vendor_name', 'purchase_date', 'purchase_cost', 'notes'])
    w.writerow(['AST-001', 'Ceiling Fan', 'Electrical', 'Boys Hostel 1', '1', 'good',
                'Usha Industries', '2024-01-15', '1200.00', 'Double blade fan'])
    w.writerow(['AST-002', 'Study Table', 'Furniture', 'Girls Hostel 1', '2', 'fair',
                'Local Vendor', '2023-06-10', '2500.00', ''])
    return response


@login_required
def asset_type_list(request):
    if request.user.role not in (User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        name = request.POST.get('name', '').strip()
        icon = request.POST.get('icon', 'bi-box-fill').strip()
        desc = request.POST.get('description', '').strip()
        if name:
            AssetType.objects.get_or_create(name=name, defaults={'icon': icon, 'description': desc})
            messages.success(request, f'Asset type "{name}" created.')
        return redirect('asset_type_list')
    types = AssetType.objects.annotate(asset_count=Count('assets'))
    return render(request, 'admin/asset_types.html', {'types': types})


# ═══════════════════════════════════════════════════════════════════════════════
# GATE PASS CATEGORY MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

@login_required
def gp_category_list(request):
    if request.user.role != User.Role.SUPER_ADMIN:
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        action = request.POST.get('action', 'create')
        name   = request.POST.get('name', '').strip()
        if not name:
            messages.error(request, 'Name required.')
            return redirect('gp_category_list')
        cat = get_object_or_404(GatePassCategory, pk=request.POST.get('pk')) if action == 'edit' else GatePassCategory(created_by=request.user)
        cat.name        = name
        cat.description = request.POST.get('description', '').strip()
        cat.start_time  = request.POST.get('start_time') or None
        cat.end_time    = request.POST.get('end_time') or None
        cat.max_hours   = int(request.POST.get('max_hours', 12) or 12)
        cat.requires_warden_approval     = 'req_warden' in request.POST
        cat.requires_admin_approval      = 'req_admin' in request.POST
        cat.requires_superadmin_approval = 'req_superadmin' in request.POST
        cat.is_emergency = 'is_emergency' in request.POST
        cat.is_active    = 'is_active' in request.POST
        cat.save()
        messages.success(request, f'Category "{cat.name}" saved.')
        return redirect('gp_category_list')
    categories = GatePassCategory.objects.annotate(pass_count=Count('gate_passes'))
    return render(request, 'admin/gp_categories.html', {'categories': categories})


@login_required
@require_POST
def gp_category_toggle(request, pk):
    if request.user.role != User.Role.SUPER_ADMIN:
        return redirect('dashboard')
    cat = get_object_or_404(GatePassCategory, pk=pk)
    cat.is_active = not cat.is_active
    cat.save()
    return redirect('gp_category_list')


# ═══════════════════════════════════════════════════════════════════════════════
# ACADEMIC CALENDAR
# ═══════════════════════════════════════════════════════════════════════════════

@login_required
def academic_calendar(request):
    if request.user.role != User.Role.SUPER_ADMIN:
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        action = request.POST.get('action', 'add')
        if action == 'delete':
            get_object_or_404(AcademicCalendarEvent, pk=request.POST.get('pk')).delete()
            messages.success(request, 'Event removed.')
            return redirect('academic_calendar')
        date_val   = request.POST.get('date', '').strip()
        event_type = request.POST.get('event_type', 'holiday')
        title      = request.POST.get('title', '').strip()
        if not date_val or not title:
            messages.error(request, 'Date and title required.')
            return redirect('academic_calendar')
        AcademicCalendarEvent.objects.update_or_create(
            date=date_val,
            defaults={'event_type': event_type, 'title': title,
                      'description': request.POST.get('description', '').strip(),
                      'marked_by': request.user}
        )
        messages.success(request, f'Event saved for {date_val}.')
        return redirect('academic_calendar')

    import calendar as cal_mod
    from datetime import date as dt
    year  = int(request.GET.get('year',  timezone.now().year))
    month = int(request.GET.get('month', timezone.now().month))
    if month < 1:  month, year = 12, year - 1
    if month > 12: month, year = 1, year + 1

    first_day = dt(year, month, 1)
    events = AcademicCalendarEvent.objects.filter(date__year=year, date__month=month)
    events_by_date = {e.date.day: e for e in events}
    cal = cal_mod.monthcalendar(year, month)

    return render(request, 'admin/academic_calendar.html', {
        'year': year, 'month': month,
        'month_name': first_day.strftime('%B'),
        'cal': cal, 'events_by_date': events_by_date,
        'recent_events': AcademicCalendarEvent.objects.order_by('-date')[:20],
        'event_types': AcademicCalendarEvent.EventType.choices,
        'today': timezone.now().date(),
        'prev_year':  year if month > 1 else year - 1,
        'prev_month': month - 1 if month > 1 else 12,
        'next_year':  year if month < 12 else year + 1,
        'next_month': month + 1 if month < 12 else 1,
    })


# ═══════════════════════════════════════════════════════════════════════════════
# LEAVE CATEGORY MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

@login_required
def leave_category_list(request):
    if request.user.role != User.Role.SUPER_ADMIN:
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    if request.method == 'POST':
        action = request.POST.get('action', 'create')
        name   = request.POST.get('name', '').strip()
        if not name:
            messages.error(request, 'Name required.')
            return redirect('leave_category_list')
        cat = get_object_or_404(LeaveCategory, pk=request.POST.get('pk')) if action == 'edit' else LeaveCategory(created_by=request.user)
        cat.name        = name
        cat.description = request.POST.get('description', '').strip()
        cat.max_days    = int(request.POST.get('max_days', 7) or 7)
        cat.start_time  = request.POST.get('start_time') or None
        cat.end_time    = request.POST.get('end_time') or None
        cat.requires_warden_approval     = 'req_warden' in request.POST
        cat.requires_admin_approval      = 'req_admin' in request.POST
        cat.requires_superadmin_approval = 'req_superadmin' in request.POST
        cat.is_active = 'is_active' in request.POST
        cat.save()
        messages.success(request, f'Leave category "{cat.name}" saved.')
        return redirect('leave_category_list')
    categories = LeaveCategory.objects.annotate(leave_count=Count('leave_applications'))
    return render(request, 'admin/leave_categories.html', {'categories': categories})


@login_required
@require_POST
def leave_category_toggle(request, pk):
    if request.user.role != User.Role.SUPER_ADMIN:
        return redirect('dashboard')
    cat = get_object_or_404(LeaveCategory, pk=pk)
    cat.is_active = not cat.is_active
    cat.save()
    return redirect('leave_category_list')


@login_required
def leave_extension_request(request, leave_pk):
    student, redir = _get_student_or_redirect(request)
    if redir: return redir
    leave = get_object_or_404(LeaveApplication, pk=leave_pk, student=student, status='approved')
    if request.method == 'POST':
        new_to_date = request.POST.get('new_to_date', '').strip()
        reason      = request.POST.get('reason', '').strip()
        if not new_to_date or not reason:
            messages.error(request, 'New date and reason required.')
            return redirect('leave_extension_request', leave_pk=leave_pk)
        from datetime import date as dt
        try:
            ndate = dt.fromisoformat(new_to_date)
        except ValueError:
            messages.error(request, 'Invalid date.')
            return redirect('leave_extension_request', leave_pk=leave_pk)
        if ndate <= leave.to_date:
            messages.error(request, 'New date must be after current end date.')
            return redirect('leave_extension_request', leave_pk=leave_pk)
        LeaveExtensionRequest.objects.create(leave=leave, new_to_date=ndate, reason=reason)
        send_notification(_staff_targets(student),
                          f'Leave Extension — {student.name}',
                          f'{student.name} requested extension until {ndate}.',
                          notif_type='leave', link=f'/leave/{leave.pk}/action/')
        messages.success(request, 'Extension request submitted.')
        return redirect('my_leaves')
    return render(request, 'student/leave_extension.html', {'leave': leave})


@login_required
@require_POST
def leave_extension_action(request, ext_pk):
    role = request.user.role
    if role not in (User.Role.WARDEN, User.Role.ADMIN, User.Role.SUPER_ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    ext    = get_object_or_404(LeaveExtensionRequest, pk=ext_pk, status='pending')
    action = request.POST.get('action', '')
    if action == 'approve':
        ext.status = LeaveExtensionRequest.Status.APPROVED
        ext.approved_by = request.user
        ext.approved_at = timezone.now()
        ext.remarks     = request.POST.get('remarks', '').strip()
        ext.save()
        ext.leave.to_date = ext.new_to_date
        ext.leave.save(update_fields=['to_date'])
        messages.success(request, f'Extension approved to {ext.new_to_date}.')
    elif action == 'reject':
        ext.status  = LeaveExtensionRequest.Status.REJECTED
        ext.remarks = request.POST.get('remarks', '').strip()
        ext.save()
        messages.info(request, 'Extension rejected.')
    return redirect('leave_list')


@login_required
@require_POST
def leave_warden_action(request, pk):
    if request.user.role not in (User.Role.WARDEN, User.Role.SUPER_ADMIN, User.Role.ADMIN):
        messages.error(request, 'Access denied.')
        return redirect('dashboard')
    leave   = get_object_or_404(LeaveApplication, pk=pk)
    action  = request.POST.get('action', '')
    remarks = request.POST.get('remarks', '').strip()
    if action == 'warden_approve':
        leave.warden_status      = 'approved'
        leave.warden_approved_by = request.user
        leave.warden_approved_at = timezone.now()
        leave.admin_remarks      = remarks
        cat = leave.category
        if not cat or not cat.requires_admin_approval:
            leave.status      = LeaveApplication.Status.APPROVED
            leave.reviewed_by = request.user
            leave.reviewed_at = timezone.now()
        leave.save()
        messages.success(request, 'Leave approved by warden.')
    elif action == 'warden_reject':
        leave.warden_status      = 'rejected'
        leave.warden_approved_by = request.user
        leave.warden_approved_at = timezone.now()
        leave.status             = LeaveApplication.Status.REJECTED
        leave.admin_remarks      = remarks
        leave.reviewed_by        = request.user
        leave.reviewed_at        = timezone.now()
        leave.save()
        messages.info(request, 'Leave rejected.')
    return redirect('leave_list')
