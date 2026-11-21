Migration preview
=================

.. automodule:: gspread.migration
   :members:
   :undoc-members:

The entry point is
:meth:`gspread.worksheet.Worksheet.migration_preview`, which builds a
:class:`~gspread.migration.MigrationPlan` without performing any write.

A plan is inert: review it, then send its payloads to
:meth:`gspread.worksheet.Worksheet.batch_update`::

    plan = worksheet.migration_preview(
        destination_title="New template",
        column_mapping={"A": 1, "C": 2, "F": 3},
        source_range="A1:F100",
    )

    destination = spreadsheet.worksheet("New template")
    for group in plan.batch_groups:
        destination.batch_update(
            group.data,
            value_input_option=group.value_input_option,
        )

When a formula cannot be remapped safely (a missing column, a split or
reversed rectangle, an open-ended reference or a dynamic reference such as
``INDIRECT`` / ``OFFSET``), :class:`~gspread.migration.MigrationPlanError`
is raised and every offending source formula cell is listed in its
``issues`` attribute.
