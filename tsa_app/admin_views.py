import os
from django.conf import settings
from django.contrib.admin.views.decorators import staff_member_required
from django.http import FileResponse, Http404
from django.shortcuts import render
from django.contrib.admin import site as admin_site
from urllib.parse import quote


@staff_member_required
def media_browser(request):
    base_path = settings.MEDIA_ROOT
    # relative_path = request.GET.get("path", "").strip("/")
    relative_path = request.GET.get("path")
    if not relative_path:
        relative_path = ""
    else:
        relative_path = relative_path.strip("/")

    abs_path = os.path.join(base_path, relative_path)

    if not abs_path.startswith(base_path):
        raise Http404("Invalid path")

    items = []
    try:
        for name in sorted(os.listdir(abs_path)):
            full_path = os.path.join(abs_path, name)
            rel_path = os.path.relpath(full_path, base_path)

            if os.path.isdir(full_path):
                items.append({
                    "name": name + "/",
                    "path": rel_path,
                    "is_dir": True,
                })
            else:
                items.append({
                    "name": name,
                    "path": rel_path,
                    "is_dir": False,
                })
    except FileNotFoundError:
        raise Http404("Directory not found")

    parent = os.path.dirname(relative_path) if relative_path else None

    context = admin_site.each_context(request)
    context.update({
        "items": items,
        "parent": parent,
        "current": relative_path,
    })
    return render(request, "admin/media_browser.html", context)


@staff_member_required
def download_file(request):
    rel_path = request.GET.get("path", "").strip("/")
    abs_path = os.path.join(settings.MEDIA_ROOT, rel_path)

    if not os.path.exists(abs_path) or not abs_path.startswith(settings.MEDIA_ROOT):
        raise Http404("File not found")

    filename = os.path.basename(abs_path)
    return FileResponse(open(abs_path, 'rb'), as_attachment=True, filename=quote(filename))
