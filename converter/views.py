import os
import json
import uuid
from pathlib import Path
from django.shortcuts import render
from django.http import JsonResponse, FileResponse, Http404
from django.conf import settings
from django.views.decorators.csrf import csrf_exempt

from core.converter import convert_ppt_to_xls
from core.models import MappingConfig
from core.ppt_extractor import extract_presentation_content


def get_demo_assets_path():
    return Path(settings.BASE_DIR) / "test_assets"


def index(request):
    """Renders the main converter interface."""
    return render(request, "converter/index.html")


def load_demo(request):
    """Loads sample budget mapping and deck info."""
    assets_dir = get_demo_assets_path()
    mapping_file = assets_dir / "budget_mapping.json"
    
    if not mapping_file.exists():
        return JsonResponse({"error": "Demo assets not found. Run generate_test_assets.py"}, status=404)
        
    with open(mapping_file, "r", encoding="utf-8") as f:
        mapping_data = json.load(f)
        
    return JsonResponse({
        "mapping_json": json.dumps(mapping_data, indent=2),
        "standard_deck": "sample_budget_deck.pptx",
        "variant_deck": "sample_budget_variant_deck.pptx"
    })


@csrf_exempt
def convert_files(request):
    """Processes PPT and JSON files, producing an Excel workbook."""
    if request.method != "POST":
        return JsonResponse({"error": "POST method required"}, status=405)

    try:
        use_demo = request.POST.get("use_demo")
        api_key = os.getenv("GEMINI_API_KEY", "")

        assets_dir = get_demo_assets_path()
        ppt_source = None
        mapping_source = None

        if use_demo:
            if use_demo == "variant":
                ppt_source = str(assets_dir / "sample_budget_variant_deck.pptx")
            else:
                ppt_source = str(assets_dir / "sample_budget_deck.pptx")
            
            custom_mapping_text = request.POST.get("mapping_text", "").strip()
            if custom_mapping_text:
                mapping_source = custom_mapping_text
            else:
                mapping_source = str(assets_dir / "budget_mapping.json")
        else:
            if "ppt_file" not in request.FILES:
                return JsonResponse({"error": "No PPT file uploaded"}, status=400)
            ppt_source = request.FILES["ppt_file"]

            if "mapping_file" in request.FILES:
                mapping_source = request.FILES["mapping_file"].read().decode("utf-8")
            elif request.POST.get("mapping_text", "").strip():
                mapping_source = request.POST.get("mapping_text").strip()
            else:
                return JsonResponse({"error": "No mapping JSON provided (upload file or paste JSON)"}, status=400)

        # Execute conversion
        override_data = None
        demo_payload = {
            "fiscal_year": "FY2025",
            "total_budget": 4500000,
            "headcount": 85,
            "departments": [
                {"dept_name": "Engineering", "planned": 2100000, "actual": 510000},
                {"dept_name": "Marketing & Sales", "planned": 1200000, "actual": 290000},
                {"dept_name": "Operations & Support", "planned": 800000, "actual": 180000},
                {"dept_name": "Human Resources", "planned": 400000, "actual": 95000}
            ]
        }

        # If API key is placeholder and running demo, use verified deterministic payload
        if (not api_key or "YOUR_GEMINI_API_KEY" in api_key) and use_demo:
            override_data = demo_payload

        try:
            wb, extracted_data, deck = convert_ppt_to_xls(
                ppt_source=ppt_source,
                mapping_source=mapping_source,
                api_key=api_key,
                extracted_data_override=override_data
            )
        except Exception as live_err:
            if use_demo:
                # In demo mode, gracefully fall back if Google API experiences transient spikes
                wb, extracted_data, deck = convert_ppt_to_xls(
                    ppt_source=ppt_source,
                    mapping_source=mapping_source,
                    api_key=api_key,
                    extracted_data_override=demo_payload
                )
                override_data = demo_payload
            else:
                raise live_err

        # Save generated Excel
        file_id = uuid.uuid4().hex
        output_filename = f"converted_{file_id}.xlsx"
        output_filepath = os.path.join(settings.MEDIA_ROOT, output_filename)
        wb.save(output_filepath)

        return JsonResponse({
            "success": True,
            "file_id": file_id,
            "download_url": f"/download/{file_id}/",
            "extracted_data": extracted_data,
            "demo_fallback_used": override_data is not None,
            "slides_count": len(deck.slides),
            "deck_preview": deck.to_markdown()[:1500] + ("..." if len(deck.to_markdown()) > 1500 else ""),
        })

    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)


def download_excel(request, file_id):
    """Serves the generated Excel file for download."""
    clean_id = "".join(c for c in file_id if c.isalnum())
    file_path = os.path.join(settings.MEDIA_ROOT, f"converted_{clean_id}.xlsx")
    
    if not os.path.exists(file_path):
        raise Http404("Excel file not found or expired.")
        
    return FileResponse(
        open(file_path, "rb"),
        as_attachment=True,
        filename="budget_planning_export.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
