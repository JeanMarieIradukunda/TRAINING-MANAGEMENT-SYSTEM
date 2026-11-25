from django.urls import path

from . import exam_builder

# Exam Builder routes. Included from training_management/urls.py so that the
# existing core/urls.py does not need to change.
urlpatterns = [
    path('exam-builder/', exam_builder.ExamBuilderView.as_view(), name='exam-builder'),
    path('api/exam-builder/ai-draft/', exam_builder.exam_ai_draft, name='exam-ai-draft'),
    path('api/exam-builder/export/docx/', exam_builder.export_exam_docx, name='exam-export-docx'),
]
